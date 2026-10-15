import { ObjectStore } from "./ObjectStore";
import { branchNameFromHead, currentRef, findRef } from "./History";

/**
 * Branch names are stored as XML attributes and are shown in the task pane, so
 * they have to be representable in both. This rejects the characters that are
 * not -- control characters, and the whitespace that XML attribute-value
 * normalisation would silently rewrite to a space -- and nothing else.
 */
const BAD_BRANCH_CHAR = new RegExp("[\\u0000-\\u001F\\u007F]");

export function validateBranchName(name: string): string {
  const trimmed = name.trim();
  if (!trimmed) throw new Error("A branch needs a name.");
  if (BAD_BRANCH_CHAR.test(trimmed)) {
    throw new Error("A branch name cannot contain control characters or line breaks.");
  }
  return trimmed;
}

export class BranchManager {
  private store: ObjectStore;

  constructor(store: ObjectStore) {
    this.store = store;
  }

  /** The current branch name (e.g. "main"). */
  async currentBranch(): Promise<string> {
    return branchNameFromHead((await this.store.load()).head);
  }

  /** All branch refs { name, tip }. */
  async listBranches(): Promise<Array<{ name: string; tip: string }>> {
    return (await this.store.load()).refs;
  }

  /** The tip commit hash for a named branch. Throws if not found. */
  async getTip(branchName: string): Promise<string> {
    const ref = findRef(await this.store.load(), branchName);
    if (!ref) throw new Error(`Branch "${branchName}" not found.`);
    return ref.tip;
  }

  /** Create a new branch pointing at the current branch's tip. */
  async createBranch(name: string): Promise<void> {
    const branch = validateBranchName(name);
    const repo = await this.store.load();
    if (repo.refs.length === 0) {
      throw new Error("There is no repository in this document yet. Init first.");
    }
    if (findRef(repo, branch)) throw new Error(`Branch "${branch}" already exists.`);
    const current = currentRef(repo);
    if (!current) {
      throw new Error(`HEAD points at "${branchNameFromHead(repo.head)}", which does not exist.`);
    }
    repo.refs.push({ name: branch, tip: current.tip });
    await this.store.save(repo);
  }

  /**
   * Switch HEAD to a different branch and return the tip to check out.
   *
   * Returning the tip is what keeps this to a single read of the document: the
   * caller used to switch, then ask for the tip, then reconstruct, and every
   * one of those re-read and re-parsed the whole repository out of the Custom
   * XML Part.
   */
  async switchBranch(name: string): Promise<string> {
    const repo = await this.store.load();
    const ref = findRef(repo, name);
    if (!ref) throw new Error(`Branch "${name}" not found.`);
    repo.head = `refs/heads/${name}`;
    await this.store.save(repo);
    return ref.tip;
  }
}
