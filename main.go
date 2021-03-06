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
var ff string
var fc string
var fk string
var fa bool

func init() {
  flag.StringVar( &fs, "s", "", "specified sheet name" )
  flag.StringVar( &ff, "f", "\t", "field separator" )
  flag.StringVar( &fc, "c", "", "cut off data" )
  flag.StringVar( &fk, "k", "", "keyword for search" )
  flag.BoolVar( &fa, "a", false, "searching all" )
}

func open( fileName string ) ( *xlsx.File, error ) {
  if fileName == "" {
    b, err := ioutil.ReadAll( os.Stdin )
    if err != nil {
      panic( err )
    }
    return xlsx.OpenBinary( b )
  }

  _, err := os.Stat( fileName )
  if err != nil {
    fmt.Println( flag.Arg( 0 ) + " is not file." )
    os.Exit( 1 )
  }
  return xlsx.OpenFile( flag.Arg( 0 ) )
}

func main() {
  flag.Parse()
  if _, err := strconv.Unquote( `"` + ff + `"` ); err != nil {
    panic( err )
  }
  ff, _ = strconv.Unquote( `"` + ff + `"` )

  if 1 < flag.NArg() {
    panic( "too many arguments." )
  }

  // Excel ファイルを読み込む
  book, err := open( flag.Arg( 0 ) )
  if err != nil {
    panic( err )
  }

  if len( fs ) == 0 && len( fc ) == 0 &&len( fk ) == 0 {
    for _, sheet := range book.Sheets {
      fmt.Println( sheet.Name )
    }
    os.Exit( 0 )
  }

  // evaluate fs
  if len( fs ) > 0 {
    _, ok := book.Sheet[ fs ]
    if !ok {
      fmt.Println( "sheet not found." )
      os.Exit( 1 )
    }
  }


  // evaluate fk, fa
  for _, sheet := range book.Sheets {
    if len( fs ) > 0 && sheet.Name != fs { continue }
    var cmin, cmax, rmin, rmax int = evaluateFc( fc, sheet.MaxCol, sheet.MaxRow )
    for r, row := range sheet.Rows {
      if r < rmin || rmax < r { continue }
      for c, cell := range row.Cells {
        if c < cmin || cmax < c { continue }
        if len( fk ) > 0 && regexp.MustCompile( fk ).MatchString( cell.String() ) {
          sn := sheet.Name + "!"
          if len( fs ) > 0 { sn = "" }
          fmt.Printf( "%s%s%d\tText=[%s]\n", sn, xlsx.ColIndexToLetters( c ), r+1, cell.String() )
          if fa == false {
            os.Exit( 0 )
          }
        } else if len( fk ) == 0 {
          fmt.Print( cell.String(), ff )
        }
      }
      if len( fk ) == 0 { fmt.Println() }
    }
  }

}

// evaluate fc
func evaluateFc( axis string, maxCol, maxRow int ) ( cmin, cmax, rmin, rmax int ) {
  cmax, rmax = maxCol, maxRow
  if len( axis ) == 0 { return }

  re := regexp.MustCompile( `^(([A-Z]+)([0-9]*))?:(([A-Z]+)([0-9]*))?$` )
  if re.MatchString( axis ) {
//fmt.Println( re.ReplaceAllString( fc, "match1=[$2] match2=[$3] match3=[$5] match4=[$6]" ) )
    if re.ReplaceAllString( axis, "$1" ) != "" {
      cmin = xlsx.ColLettersToIndex( re.ReplaceAllString( axis, "$2" ) )
      rmin, _ = strconv.Atoi( re.ReplaceAllString( axis, "$3" ) )
      rmin--
    }
    if re.ReplaceAllString( axis, "$4" ) != "" {
      cmax = xlsx.ColLettersToIndex( re.ReplaceAllString( axis, "$5" ) )
      rmax, _ = strconv.Atoi( re.ReplaceAllString( axis, "$6" ) )
      rmax--
    }
  } else {
    fmt.Println( "invalid axis." )
    os.Exit( 1 )
  }
  return
}

