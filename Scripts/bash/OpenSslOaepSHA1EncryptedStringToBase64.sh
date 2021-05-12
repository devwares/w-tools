#!/bin/bash
# Check syntax
if [ -z "$1" ]; then
  echo "Syntax : $0 [plain text file] [publickey file] [[outfile path]]"
  echo "			public key file must be at .cer format"
  exit
fi

# Check file paths
notfound=0
for fichier in "$1" "$2"
do
  if [ ! -f "$fichier" ]; then
    echo "file not found : $fichier"
    notfound=1
  fi
done
if [ $notfound = 1 ]
then
  exit
fi

# Execution
INFILE=$1
PUBKEY=$2
if [ -z "$3" ]; then
  openssl pkeyutl -in $1 -encrypt -pubin -inkey $2 -pkeyopt rsa_padding_mode:oaep -pkeyopt rsa_oaep_md:sha1 -pkeyopt rsa_mgf1_md:sha1 | base64 -w0
else
  openssl pkeyutl -in $1 -encrypt -pubin -inkey $2 -pkeyopt rsa_padding_mode:oaep -pkeyopt rsa_oaep_md:sha1 -pkeyopt rsa_mgf1_md:sha1 | base64 -w0 > $3
fi