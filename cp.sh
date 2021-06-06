#!/bin/sh
mkdir -p data
for i in $(find $1 -name \*.xlsx -o -iname \*.xls); do cp $i data; done 
