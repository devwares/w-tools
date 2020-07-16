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
