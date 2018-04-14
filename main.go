package main

import (
  "flag"
  "fmt"
  "os"
  "regexp"

  "github.com/360EntSecGroup-Skylar/excelize"
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
  book, err := excelize.OpenReader( os.Stdin )
  if err != nil {
    panic( err )
  }
//  fmt.Printf( "file opened.\n" )

  if len( fs ) == 0 {
    sm := book.GetSheetMap()
    for k, v := range sm {
      fmt.Printf( "%d:\t%s\n", k, v )
    }
    os.Exit( 0 )
  }

  if book.GetSheetVisible( fs ) == false {
    fmt.Print( "sheet not found.\n" )
    os.Exit( 1 )
  }

  for i, row := range book.GetRows( fs ) {
    for j, cell := range row {
      if len( fk ) > 0 && regexp.MustCompile( fk ).MatchString( cell ) {
        fmt.Printf( "row=%d cell=%d Text=[%s]\n", i, j, cell )
        if fa == false {
          os.Exit( 0 )
        }
      } else if len( fk ) == 0 {
        fmt.Print( cell, "\t" )
      }
    }
    if len( fk ) == 0 { fmt.Println() }
  }

}

