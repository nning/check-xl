#!/bin/sh
mkdir -p data
for i in $(find ~/sshfs/mama-sda3-photorec-zip -name \*.xlsx); do cp $i data; done 
