// Where the pictures of a document are read from. A docx keeps them as parts
// of its package; a binary doc has no package, so its reader puts them into
// the model.

import type { DocModel } from './types';

export interface PictureSource {
  readPart(part: string): Promise<Uint8Array>;
}

/** The source the view loads pictures from: the model's own pictures when it carries them, else the package. */
export function documentPictures(doc: DocModel, pkg?: PictureSource): PictureSource | undefined {
  const media = doc.media;
  if (!media) return pkg;
  return {
    readPart(part: string): Promise<Uint8Array> {
      const found = media.get(part);
      return found ? Promise.resolve(found.bytes) : Promise.reject(new Error(`No such picture: ${part}`));
    },
  };
}
