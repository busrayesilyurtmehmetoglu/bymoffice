// What the Save button of a toolbar runs and shows. The toolbars know nothing
// of where a file is saved to: whoever holds the editor (mount(), the bridge)
// hands this in, and says when what it shows has changed.

export interface SaveCommand {
  /** Starts saving; what comes of it is told through `dirty` and `saving`. */
  run(): void;
  /** Whether there are changes that have not been saved. */
  readonly dirty: boolean;
  /** Whether a save is under way. */
  readonly saving: boolean;
  /** Calls `listener` whenever `dirty` or `saving` changed; the result ends that. */
  watch(listener: () => void): () => void;
}
