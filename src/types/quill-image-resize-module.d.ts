declare module 'quill-image-resize-module' {
  import type QuillType from 'quill';

  export interface IQuillImageResizeOptions {
    modules?: string[];
    displayStyles?: Record<string, string>;
    parchment?: any;
  }

  export default class ImageResize {
    constructor(quill: QuillType, options?: IQuillImageResizeOptions);
  }
}
