package main

import (
  "flag"
  "fmt"
  "io/ioutil"
  "os"
  "regexp"
  "strconv"

  "github.com/tealeg/xlsx"
)

var fs string
var fc string
var fk string
var fa bool

func init() {
  flag.StringVar( &fs, "s", "", "specified sheet name" )
  flag.StringVar( &fc, "c", "", "cut off data" )
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

  // evaluate fs
  sheet, ok := book.Sheet[ fs ]
  if !ok {
    fmt.Println( "sheet not found." )
    os.Exit( 1 )
  }

  // evaluate fc
  var cmin, cmax, rmin, rmax int = 0, sheet.MaxRow, 0, sheet.MaxCol
  if len( fc ) > 0 {
    re := regexp.MustCompile( `^(([A-Z]+)([0-9]*))?:(([A-Z]+)([0-9]*))?$` )
    if re.MatchString( fc ) {
//fmt.Println( re.ReplaceAllString( fc, "match1=[$2] match2=[$3] match3=[$5] match4=[$6]" ) )
      if re.ReplaceAllString( fc, "$1" ) != "" {
        cmin = xlsx.ColLettersToIndex( re.ReplaceAllString( fc, "$2" ) )
        rmin, _ = strconv.Atoi( re.ReplaceAllString( fc, "$3" ) )
        rmin--
      }
      if re.ReplaceAllString( fc, "$4" ) != "" {
        cmax = xlsx.ColLettersToIndex( re.ReplaceAllString( fc, "$5" ) )
        rmax, _ = strconv.Atoi( re.ReplaceAllString( fc, "$6" ) )
        rmax--
      }
    } else {
      fmt.Println( "invalid axis." )
      os.Exit( 1 )
    }
  }
//fmt.Printf( "cmin=%d cmax=%d rmin=%d rmax=%d\n", cmin, cmax, rmin, rmax )

  // evaluate fk, fa
  for r, row := range sheet.Rows {
    if r < rmin || rmax < r { continue }
    for c, cell := range row.Cells {
      if c < cmin || cmax < c { continue }
      if len( fk ) > 0 && regexp.MustCompile( fk ).MatchString( cell.String() ) {
        fmt.Printf( "%s%d\tText=[%s]\n", xlsx.ColIndexToLetters( c ), r+1, cell.String() )
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

