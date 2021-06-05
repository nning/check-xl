# check-xlsx

Small utility to parse several xlsx files and output a list of contained sheet names.
Can be used to get an overview over files recovered by photorec without
structure and original file names.

## Example

    go build
    ./cp.sh ~/sshfs/sda3-photorec-zip
    ./check-xlsx -s data