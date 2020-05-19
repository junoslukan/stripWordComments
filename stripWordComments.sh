#!/bin/bash
# A script to replace the name and/or date of comments in a Microsoft Word .docx file
# Use -f 'file name' -a 'new author name' and -d to remove date.

usage="$(basename "$0") [-h] [-f 'filename' -a 'authorname' -d]
-- A script to replace the name and/or date of comments in a Microsoft Word .docx file

where:
    -h  show this help text
    -f specify the name .docx file
    -a specify the new author name
    -d specify if date of comments should be removed (ignored)"

mkdir ./temp;
f_present=0;

while getopts "f::a:d" opt; do
  case $opt in
    h)
      echo "$usage" >&2
      exit
      ;;
    f)
      FILENAME=$OPTARG
      unzip $FILENAME -d ./temp
      echo "Copied and extracted the file." >&2
      f_present=1;
      ;;
    a)
      if [ $f_present -eq 0 ]; then
        rm -r ./temp;
        echo "ERROR: Missing the -f argument." >&2
        echo "$usage" >&2
        exit 1
      fi
      AUTHOR=$OPTARG
      sed -i -e 's/w:author="Author"/w:author="'"$AUTHOR"'"/g' ./temp/word/comments.xml
      sed -i -e 's/w:author="Author"/w:author="'"$AUTHOR"'"/g' ./temp/word/document.xml
      echo "Replaced the author name." >&2
      ;;
    d)
      if [ $f_present -eq 0 ]; then
        rm -r ./temp;
        echo "ERROR: Missing the -f argument." >&2
        echo "$usage" >&2
        exit 1
      fi
      sed -i -e 's/w:date/w:ignore/g' ./temp/word/comments.xml
      sed -i -e 's/w:date/w:ignore/g' ./temp/word/document.xml
      echo "Removed the date."
      ;;
    \?) printf "illegal option: %s\n" "$OPTARG" >&2
      echo "$usage" >&2
      exit 1
      ;;
  esac
done

FILENAMENEW="${FILENAME%.docx}_new.docx"
echo $FILENAMENEW;

cd temp;
zip -r ../$FILENAMENEW *;
cd ..;
rm -r ./temp;
exit 0;
