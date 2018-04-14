package main

import (
  "flag"
  "fmt"
  "io/ioutil"
  "os"
  "regexp"

  "github.com/tealeg/xlsx"
)

var fk string
var fa bool

func init() {
  flag.StringVar( &fk, "k", "", "keyword for search" )
  flag.BoolVar( &fa, "a", false, "searching all" )
}

func main() {
  flag.Parse()

  // Excel ファイルを読み込む
  b, err := ioutil.ReadAll( os.Stdin )
  if err != nil {
    panic( err )
  }

  book, err := xlsx.OpenBinary( b )
  if err != nil {
    panic( err )
  }
//  fmt.Printf( "file opened.\n" )

  for i, sheet := range book.Sheets {
    for j, row := range sheet.Rows {
      if len( fk ) == 0 { fmt.Printf( "(%s):\t", sheet.Name ) }
      for k, cell := range row.Cells {
        if len( fk ) > 0 && regexp.MustCompile( fk ).MatchString( cell.String() ) {
          fmt.Printf( "sheet=%d:(%s) row=%d cell=%d Text=[%s]\n", i, sheet.Name, j, k, cell.String() )
          if fa == false {
            os.Exit( 0 )
          }
        } else if len( fk ) == 0 {
          if k > 0 { fmt.Print( "\t" ) }
          fmt.Print( cell.String() )
        }
      }
      if len( fk ) == 0 { fmt.Print( "\n" ) }
    }
  }

}

