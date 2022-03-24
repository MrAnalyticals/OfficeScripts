function main(workbook: ExcelScript.Workbook) {
	let GreatArt = workbook.getWorksheet('GreatArt')
	let GreatArtRange = GreatArt.getRange('a1:m30')
	let hexStr1: string
	let hexStr2: string
	let hexStr3: string
	let hexconcat:string
	let GreatArtVals1 = GreatArtRange.getValues()

	GreatArtVals1.forEach((rowItem, rowIndex) => {
		GreatArtVals1[rowIndex].forEach((columnItem, columnIndex) => {
			hexStr1 = toHex(getRandomInt(256))
			if (hexStr1.length == 1){
				hexStr1 = '0'+hexStr1}
			hexStr2 = toHex(getRandomInt(256))
			if (hexStr2.length == 1) {
				hexStr2 = '0' + hexStr2
			}
			hexStr3 = toHex(getRandomInt(256))
			if (hexStr3.length == 1) {
				hexStr3 = '0' + hexStr3
			}
			hexconcat = hexStr1.concat(hexStr2).concat(hexStr3)
			console.log(hexconcat)
			GreatArt.getCell(rowIndex, columnIndex).getFormat().getFill().setColor(hexconcat)
		})
	})


	return
}

/**
 * Convert a Number to Hexadecimal
 * Given an integer, write an algorithm to
 * convert it to hexadecimal. For negative
 * integer, two’s complement method is used.
 *
 * Time Complexity: O(log(n))
 * Space Complexity: O(log(n))
 *
 * toHex(26) // "1a"
 * toHex(4)  // "4"
 * toHex(-1) // "ffffffff"
 * Source: https://codybonney.com/leetcode-convert-a-number-to-hexadecimal-solution-using-typescript/
 */
	function toHex(num: number): string {
		const map = "0123456789abcdef";
		let hex = num === 0 ? "0" : "";
		while (num !== 0) {
			hex = map[num & 15] + hex;
			num = num >>> 4;
		}
		return hex;
	}

	function getRandomInt(max:number) {
		return Math.floor(Math.random() * max);
	}
