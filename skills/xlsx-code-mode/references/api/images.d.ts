// Image APIs
// Witan xlsx-code-mode exec API reference. All functions are async and take `wb` first.

type ImageFormat="png"|"jpeg"
interface ImagePositionAnchor {cell:string;xOffsetPts?:number;yOffsetPts?:number}
interface ImagePositionInput {from:ImagePositionAnchor;to:ImagePositionAnchor}
interface ImagePosition extends ImagePositionInput {sheet?:string}
interface ImageSource {base64:string}
interface ImageSpec {
	name:string;
	position:ImagePositionInput;
	source:ImageSource;
	format?:ImageFormat;
	altText?:string|null;
	altTextTitle?:string|null;
	preserveAspectRatio?:boolean;
}
interface ImageUpdate {
	name?:string;
	position?:ImagePositionInput;
	source?:ImageSource;
	format?:ImageFormat;
	altText?:string|null;
	altTextTitle?:string|null;
	preserveAspectRatio?:boolean;
}
interface ImageInfo {
	id?:number;
	sheet:string;
	name:string;
	position:ImagePosition;
	format?:ImageFormat;
	widthPts?:number;
	heightPts?:number;
	naturalWidthPx?:number;
	naturalHeightPx?:number;
	altText?:string|null;
	altTextTitle?:string|null;
}
declare function listImages(wb,options?:{ sheet?:string }):Promise<ImageInfo[]>;
declare function getImage(wb,sheet:string,selector:{ name?:string;id?:number }):Promise<ImageInfo>;
declare function addImage(wb,sheet:string,image:ImageSpec):Promise<ImageInfo>;
declare function setImage(wb,sheet:string,selector:{ name?:string;id?:number },image:ImageUpdate):Promise<ImageInfo>;
declare function deleteImage(wb,sheet:string,selector:{ name?:string;id?:number }):Promise<void>;
