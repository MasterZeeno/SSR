#!/usr/bin/env bash

set -euo pipefail

# --- Config ---
SSR_FILE="SSR.xlsx"
MAX_RETRY=3
IS_TERMUX=1
DRY_RUN=0
SENT_REPORTS="SENT_REPORTS" #Added for clarity and potential portability

# --- Helper Functions ---
# Simplified spacer function
spacer() { printf '%*s' "$1"; }

#Improved error handling and messaging
handle_error() {
  local msg="$1"
  local details="${2:-}"
  local extra="${3:-}"
  send_msg 1 "$msg" "$details" "$extra" >&2
  exit 1
}

# --- Main Functions ---
install_pkgs() {
  local sudo='' pkgs=(gpg mutt msmtp python3)
  ((IS_TERMUX)) || {
    [[ $(id -u) -eq 0 ]] || sudo='sudo'
    pkgs+=(python3-pip)
  }

  for pkg in "${pkgs[@]}" xlsx2csv; do
    until command -v "$pkg" &>/dev/null; do
      if (( ++retry > MAX_RETRY )); then
        handle_error "Failed to install '$pkg'" "Please check internet connection"
      fi
      [[ -z "$sudo" ]] || $sudo apt-get update -qq
      $sudo apt-get install -yq --no-install-recommends "$pkg" || sleep 1
    done
    retry=0 # Reset retry counter for each package
  done

  local pypath="$(python3 -m site --user-base)/bin"
  [[ "$PATH" != *"$pypath"* ]] && export PATH="$pypath:$PATH"
}


setup_email() {
  source gmail_data # Assumes gmail_data exists and is correctly sourced

  #Simplified GPG command
  local gpg_cmd="gpg --quiet --batch --yes --pinentry-mode loopback --passphrase '$pass_phrase'"

  local trust_file="${PREFIX:-}/etc/$(if [[ "$IS_TERMUX" -eq 1 ]]; then echo 'tls/cert.pem'; else echo 'ssl/certs/ca-certificates.crt'; fi)"

  # Function to create files, simplified
  create_file() {
    local file="$1"
    [[ -f "$file" ]] && rm -f "$file" #Removed redundant INIT check.

    if [[ ! -s "$file" ]]; then
      if [[ "$file" == *".gpg" ]]; then
        echo "$gmail_pass" | "$gpg_cmd" --symmetric -o "$file"
      elif [[ "$file" == *".muttrc" ]]; then
        #Simplified muttrc creation
        cat <<EOF > "$file"
set sendmail="$(command -v msmtp)"
set content_type = "text/html"
set use_from = yes
set realname = "$real_name"
set from = "$sender_email"
set envelope_from = yes
EOF
      elif [[ "$file" == *".msmtprc" ]]; then
        #Simplified msmtprc creation
        cat <<EOF > "$file"
defaults
auth on
tls on
tls_trust_file "$trust_file"
account gmail
host smtp.gmail.com
port 587
from "$sender_email"
user "$sender_email"
passwordeval "$gpg_cmd --decrypt $gpass_file"
logfile "$msmtp_log"
account default: gmail
EOF
      fi
      [[ -f "$file" ]] && chmod 600 "$file"
    fi
  }

  #Simplified file creation calls
  create_file "$HOME/.gpass.gpg"
  create_file "$HOME/.muttrc"
  create_file "$HOME/.msmtprc"
}

send_email() {
  setup_email

  #Simplified recipient handling
  local recipients=( "jojofundales@hcc.com.ph" "jojofundales@yahoo.com" )
  local copies=( "arch_rbporral@yahoo.com" "glachel.arao@yahoo.com" "rbzden@yahoo.com" "aljonporcalla@gmail.com" "eduardo111680@gmail.com" )

  #Simplified email construction
  local subject="NSB-P2 SSR as of $REPORT_DATE"
  local body_file="body.html"
  local cmd="mutt -s \"$subject\" -c \"$(sed 's/\s/, /g' <<<"${copies[*]}")\" -a \"$SSR_FILE\" -- \"${recipients[@]}\" <\"$body_file\""

  if ((DRY_RUN)); then
      echo "Dry run: Would have executed: $cmd"
  else
      if $cmd &>/dev/null; then
          send_msg 'Email successfully sent!'
      else
          handle_error 'Failed to send email!' 'Please check internet connection'
      fi
  fi

  [[ "$NOT_REPORTED" -eq 1 ]] && echo "$REPORT_DATE" >> "$SENT_REPORTS"
}


# --- Prechecks ---
while (( $# )); do
  case "$1" in
    *.xlsx) SSR_FILE="$1";;
    -d | --dry-run) DRY_RUN=1;;
    -i | --init) INIT=1;;
    *) break;;
  esac
  shift
done

#Improved error handling
if [[ ! -f "$SSR_FILE" ]]; then
  handle_error "Invalid excel file: '$SSR_FILE'" 'File does not exist or is corrupted'
fi

#Simplified Termux check
[[ ${PREFIX:-} =~ com.termux ]] || IS_TERMUX=0

# --- Install pkgs ---
install_pkgs

# --- Extract value ---
REPORT_DATE=$(if ((DRY_RUN)); then echo 'July 7-13, 2025'; else xlsx2csv -a --exclude_hidden_sheets "$SSR_FILE" | grep -Eo '^[-]+.*' | tail -n1 | cut -d ' ' -f4- | perl -pe 's/\b(\w+)\s+(\d+)-\1\s+(\d+),/$1 $2-$3,/;s/\b(\w)/\u$1/g' | sed -E 's/\s+$//'; fi)

# Simplified sent report check
[[ -f "$SENT_REPORTS" && grep -q "$REPORT_DATE" "$SENT_REPORTS" ]] && NOT_REPORTED=0

((DRY_RUN + NOT_REPORTED)) && send_email

