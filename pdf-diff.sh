#!/bin/bash
set -e
pdftotext "test/test-ExpandAnimations.pdf"
pdftotext "test/test-ExpandAnimationsReference.pdf"
set +e
diff_res=$(diff test/test-ExpandAnimations.txt test/test-ExpandAnimationsReference.txt)
diff_return=$?
if [ $diff_return -eq 0 ]; then
    echo "The content of the pdfs is equal!"
    exit 0
else
    echo "The content of the pdfs is different!"
    echo "Result of diff:"
    echo "$diff_res"
    exit 1
fi
