package main

import (
  "flag"
  "fmt"
  "io/ioutil"
  "os"
  "regexp"

  "github.com/tealeg/xlsx"
)

var fs string
var fk string
var fa bool

func init() {
  flag.StringVar( &fs, "s", "", "specified sheet name" )
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

  if len( fs ) == 0 {
    for _, sheet := range book.Sheets {
      fmt.Println( sheet.Name )
    }
    os.Exit( 0 )
  }

  if _, ok := book.Sheet[ fs ]; !ok {
    fmt.Println( "sheet not found." )
    os.Exit( 1 )
  }

  for i, row := range book.Sheet[ fs ].Rows {
    for j, cell := range row.Cells {
      if len( fk ) > 0 && regexp.MustCompile( fk ).MatchString( cell.String() ) {
        fmt.Printf( "%s%d\tText=[%s]\n", xlsx.ColIndexToLetters( j ), i, cell.String() )
        if fa == false {
          os.Exit( 0 )
        }
      } else if len( fk ) == 0 {
        fmt.Print( cell.String(), "\t" )
      }
    }
    if len( fk ) == 0 { fmt.Println() }
  }

}

