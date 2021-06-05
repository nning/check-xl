#!/bin/sh
mkdir -p data
for i in $(find $1 -name \*.xlsx); do cp $i data; done 
