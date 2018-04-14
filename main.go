package main

import (
  "flag"
  "fmt"
  "io/ioutil"
  "os"
  "strings"

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

  for _, sheet := range book.Sheets {
    for j, row := range sheet.Rows {
      for k, cell := range row.Cells {
        if strings.Contains( cell.String(), fk ) {
          fmt.Printf( "\trow=%d cell=%d Text=[%s]\n", j, k, cell.String() )
          if fa == false {
            os.Exit( 0 )
          }
        }
      }
    }
  }

}

