export function extractText(node: any, path: string[]) {
	if (path.constructor !== Array) {
		throw Error("Error of path type! path is not array.");
	}

	if (node === undefined) {
		return undefined;
	}

	let l = path.length;
	for (let i = 0; i < l; i++) {
		node = node[path[i]];
		if (node === undefined) {
			return undefined;
		}
	}

	return node;
}

export function img2Base64(data: any) {
	let base64 = '';
	let encodings = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
	let bytes = new Uint8Array(data);
	let byteLength = bytes.byteLength;
	let byteRemainder = byteLength % 3;
	let mainLength = byteLength - byteRemainder;

	let a, b, c, d;
	let chunk;

	for (let i = 0; i < mainLength; i = i + 3) {
		chunk = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];
		a = (chunk & 16515072) >> 18;
		b = (chunk & 258048) >> 12;
		c = (chunk & 4032) >> 6;
		d = chunk & 63;
		base64 += encodings[a] + encodings[b] + encodings[c] + encodings[d];
	}

	if (byteRemainder == 1) {
		chunk = bytes[mainLength];
		a = (chunk & 252) >> 2;
		b = (chunk & 3) << 4;
		base64 += encodings[a] + encodings[b] + '==';
	} else if (byteRemainder == 2) {
		chunk = (bytes[mainLength] << 8) | bytes[mainLength + 1];
		a = (chunk & 64512) >> 10;
		b = (chunk & 1008) >> 4;
		c = (chunk & 15) << 2;
		base64 += encodings[a] + encodings[b] + encodings[c] + '=';
	}

	return base64;
}

export function toBase64ImgLink(mimeType: string, buff: ArrayBuffer) {
	return `data:${mimeType};base64,${img2Base64(buff)}`
}

export function getImgMimeType(imgName: string) {
	let imgFileExt = extractFileExtension(imgName).toLowerCase();

	let mimeType
	switch (imgFileExt) {
		case "jpg":
		case "jpeg":
			mimeType = "image/jpeg";
			break;
		case "png":
			mimeType = "image/png";
			break;
		case "gif":
			mimeType = "image/gif";
			break;
		case "emf": // Not native support
			mimeType = "image/x-emf";
			break;
		case "wmf": // Not native support
			mimeType = "image/x-wmf";
			break;
		case "tiff":
			mimeType = "image/tiff";
			break;
		default:
			mimeType = "image/*";
	}

	return mimeType
}

export function computePixel(emus: string): number {
	return Math.round(parseInt(emus) * 96 / 914400)
}

export function extractFileExtension(filename: string) {
	return filename.substr((~-filename.lastIndexOf(".") >>> 0) + 2);
}

export function getSchemeColorFromTheme(theme: any, schemeClr: string) {
	switch (schemeClr) {
		case "tx1": schemeClr = "a:dk1"; break;
		case "tx2": schemeClr = "a:dk2"; break;
		case "bg1": schemeClr = "a:lt1"; break;
		case "bg2": schemeClr = "a:lt2"; break;
	}
	let refNode = extractText(theme, ["a:theme", "a:themeElements", "a:clrScheme", schemeClr]);
	let color = extractText(refNode, ["a:srgbClr", "attrs", "val"]);
	if (color === undefined) {
		color = extractText(refNode, ["a:sysClr", "attrs", "lastClr"]);
	}

	return color;
}

export function printObj(obj: any) {
	console.log(JSON.stringify(obj, null, 2))
}
