#!/bin/bash
set -e
pdftotext "test/test-ExpandAnimations.pdf"
pdftotext "test/test-ExpandAnimationsReference.pdf"
diff_files=$(diff test/test-ExpandAnimations.txt test/test-ExpandAnimationsReference.txt)
if [ "$diff_files" = "" ]; then
    echo "The content of the pdfs is equal!"
    exit 0
else
    echo "The content of the pdfs is different!"
    echo "$diff_files"
    exit 1
fi


