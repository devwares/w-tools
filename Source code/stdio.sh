#!/bin/bash

# Function to wait for reply
pressakey() {
    while [ true ] ; do
    read -t 3 -n 1
    if [ $? = 0 ] ; then
    break ;
    fi
    done
}

filenamewithoutext() {
    fnwoefullfilename=$(basename -- "$1")
    fnwoefilename="${fnwoefullfilename%.*}"
    echo "$fnwoefilename"
}

fileextension() {
    fextfullfilename=$(basename -- "$1")
    fextfullfileext="${fextfullfilename##*.}"
    echo "$fextfullfileext"
}
