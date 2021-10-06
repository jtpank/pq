#!/bin/bash
VARA=$(ls ../downloads/ | head -1)
mv ../downloads/"$VARA" ../downloads/exportRecent.csv
echo "complete"
