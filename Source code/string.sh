#!/bin/bash

# Convert $1 to lowercase
lowercase(){
    echo "$1" | sed "y/ABCDEFGHIJKLMNOPQRSTUVWXYZ/abcdefghijklmnopqrstuvwxyz/"
}

# Convert $1 to uppercase
uppercase(){
    echo "$1" | sed "y/abcdefghijklmnopqrstuvwxyz/ABCDEFGHIJKLMNOPQRSTUVWXYZ/"
}

# Replaces values from array '$confs' in file '$1'. Usage : 
# declare -A confs
# confs=(
#	[%%string1%%]="My Str 1"
#	[%%string2%%]="My str 2"
# )
# strreplaceinfile "/home/joe/file.txt"
strreplaceinfile() {

    if [ -f "$1" ]; then

        # Loop the config array
        for i in "${!confs[@]}"
        do
            search=$i
            replace=${confs[$i]}
            sed -i "s^${search}^${replace}^g" $1
        done

    else

      echo "file not found : $1"
      exit

    fi

}

