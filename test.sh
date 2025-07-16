#!/usr/bin/env bash

# app_passwd='qdqgrnajewhizoox'
app_passwd=546609529
encrypted_pass='VmpGU1IyRXhWWGxXYTJScFRUTkNWVmx0ZUdGWlZscHhWR3RPYWsxWVFrWlZNakExWVd4SmVGZHFRbFZOVjJob1dXdGFSMWRGT1VWaVJWSmhaV3BCTlZFeVl6bFFVVzg5Q2c9PQo='
max_loop=6

encrypt_pass() {
  local app_passwd_tmp1="$app_passwd"
  local app_passwd_tmp2=''
  for i in $(seq 1 $max_loop); do
    app_passwd_tmp2=$(echo "$app_passwd_tmp1" | base64 -w 0)
    app_passwd_tmp1="$app_passwd_tmp2"
  done
  echo "$app_passwd_tmp1"
}

decrypt_pass() {
  local app_passwd_tmp1="$encrypted_pass"
  local app_passwd_tmp2=''
  for i in $(seq 1 $max_loop); do
    app_passwd_tmp2=$(echo "$app_passwd_tmp1" | base64 -d -i)
    app_passwd_tmp1="$app_passwd_tmp2"
  done
  echo "$app_passwd_tmp1"
}
  
# case "$1" in
  # -e) encrypt_pass ;;
  # -d) decrypt_pass ;;
  # *) exit 1 ;;
# esac

yawa="$(p='Vm10a05GVXhWbkpOV0VwUFYwVmFVMVpyV21GVlJscHlWbTVLVGxaclZqVlhXSEJ6VlZaV1dFOUViRVJhZWpBNVEyYzlQUW89Cg==';for i in $(seq 1 6);do p=$(echo "$p" | base64 -di);done;echo "$p")"

echo "$yawa"