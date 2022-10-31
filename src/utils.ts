export function pad(n: string | number, width: number, z?: string) {
  z = z || '0'
  n = n + ''
  return n.length >= width ? n : new Array(width - n.length + 1).join(z) + n;
}

export function getLastRowSpecial(range: any[][]){
  let rowNum = 0
  let blank = false
  for(var row = 0; row < range.length; row++){
 
    if(range[row][0] === "" && !blank){
      rowNum = row;
      blank = true;
    }else if(range[row][0] !== ""){
      blank = false;
    }
  }
  return rowNum
}

export function getNamedOptions(
  options: GoogleAppsScript.Spreadsheet.NamedRange[], 
  name: string
): string[] {
  return options
  .filter((type) => type.getName() == name)[0]
  .getRange()
  .getValues() as unknown as string[]


}