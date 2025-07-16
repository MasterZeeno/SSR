#!/usr/bin/env bash

set -euo pipefail

# REPORT_DATE='July 7-13, 2025'
# sps='        '
# recipients=("jojofundales@"{"hcc.com.ph","yahoo.com"})
# copies=({"arch_rbporral","glachel.arao","rbzden"}"@yahoo.com"
        # {"aljonporcalla","eduardo111680"}"@gmail.com")
            
# eval "$(echo '${sps}' |
        # sed -E 'p;h;G;G;G;s/\n/ /g;s/ /:/;s/(([^}]+\}){1})([^}]+)/\1\3:5/
          # s/(.*) (.*)$/\1\2/;s/^/To/;s|\}$|:2&/g"|;s| |$(sed "s/ /,\\\\n|
          # s/$/ <<<"${recipients[*]}")/;p;s/To/CC/;s/recipients/copies/' |
        # sed -E '/^[^A-Z]/{s/.$/:5&/;h;G;s/\n/:/;s/^/Subject/;s/$/SSR: $REPORT_DATE/
          # p;s/5/2/;s/Subject/Date/;s/SSR.*/$(date +"%B %-d, %Y, %-l:%M %p")/}' |
        # sed -E 's/.*/ "&"/;1s/^/details=(\n/;$s/$/\n)/')"
          
# printf '%s\n' "${details[@]}"
# exit

#   sed -E 'p;h;G;G;G;s/\n/ /g;s/ /:/;s/(([^}]+\}){1})([^}]+)/\1\3:3/
            # s/(.*) (.*)$/\1\2/;s/^/To/;s|$|/g"|;s| |$(sed "s/ /,\\\\n|
            # s/$/ <<<"${recipients[*]}")/;p;s/To/CC/;s/recipients/copies/' |

# --- Config ---
SSR_FILE="SSR.xlsx"
MAX_RETRY=3
IS_TERMUX=1
NOT_REPORTED=1
DRY_RUN=0
INIT=0

spacer() {
  local num="${1:-1}"
  
  case "$num" in
    [1-9]|[1-9][0-9]) ;;
    *) num=1 ;;
  esac
  
  printf '%*s' "$num"
}

send_msg() {
  local ic='C' wht='200;250;220' ylw='23;32;0'
  local prm='15;255;80' fnt='69;96;0'

  case "$1" in
    0|1)
      if (($1)); then
        wht='250;200;220' ylw='255;253;208'
        prm='220;20;60' fnt="$wht" ic='D'
      fi
      shift || return
      ;;
  esac
  
  local -a msgs=("$@")
  (($#)) || return
  local topsp=$(spacer 2) botsp=$(spacer 4)
  if [[ -t 1 ]]; then
    clear
    eval "$({ printf '%s=\\e[0%s\n' esc '' reset m \
                  prm "$prm" ylw "$ylw" fnt "$fnt" wht "$wht" |
                sed -E '/^[^er]/s/(e\[0)(.*)/\1;38;2;\2m/'
              eval "$(printf 'prm=%s\n' ylw fnt | sed -E 'p;s/(.*)=(.*)/\2=\1/' |
                sed -E -e 's|(.)(.*)=(.)(.*)|\1\3=\${esc};%s;\${\3\4:4:-1};4\${\1\2:5}|' \
                       -e "s/.*/printf '\%s&\\\n' n 0 b 1 d 2 i 3/")"
              printf '${%s}\\UE0C\n' ylw prm npy npf |
                sed -E 's/(\$\{)(.*)(\}.*)/l_\2=\1\2\32/;p;s/^l(.*)2/r\10/'
              printf 'icon=${l_ylw}${nyp} \\UF00%s ${r_npy}\n' "$ic"
            } | sed -E "s/^/local /;s/=(.*)$/=\$(printf \"\1\")/")"
    printf "%s\n" "${msgs[@]}" | sed -E \
        "2,\${s/.*/\n${botsp}&/
          s/([[:alpha:] ]+):/${wht/0;/1;}\1${wht}:/;t;s/.*/${wht}&/}
        1{s/.*/\n${topsp}${icon}${npf}${topsp}&${botsp}${r_prm}/
          s/(excel|success.*\s)/${ipy}\1${npf}/Ig
          s/'.*'/${ipy}&${npf}/};s/$/${reset}/"
  else
    printf "%s\n" "${msgs[@]}"
  fi
  
  return 0
}

abort() {
  send_msg 1 "$@" >&2
  exit 1
}

install_fail() {
  abort "Failed to install '$1'" \
    'Please check internet connection'
}

install_pkgs() {
  local sudo='' init=0 pkgs=(gpg mutt msmtp python3)
  
  ((IS_TERMUX)) || {
    [[ $(id -u) -eq 0 ]] || sudo='sudo'
    pkgs+=(python3-pip)
  }

  for pkg in "${pkgs[@]}" xlsx2csv; do
    local retry=0; until command -v "$pkg" &>/dev/null; do
      ((retry++>MAX_RETRY)) && {
        abort "Failed to install '$pkg'" \
          'Please check internet connection'
      }
      
      ((init)) || {
        $sudo apt-get update -qq
        init=1
      }
      
      { case "$pkg" in
          xlsx2csv) $sudo pip3 install -q "$pkg" ;;
          *) $sudo env DEBIAN_FRONTEND=noninteractive \
             apt-get install -yq --no-install-recommends "$pkg" ;;
        esac
      } &>/dev/null || sleep 1; done
  done
  
  local pypath="$(python3 -m site --user-base)/bin"
  grep -q "$pypath" <<<"$PATH" || export PATH="$pypath:$PATH"
}

setup_email() {
  source gmail_data
  
  local -a gpg_cmds=(
    gpg --quiet --batch --yes --pinentry-mode
    loopback --passphrase "$pass_phrase"
  )
  
  local trust_file="${PREFIX:-}/etc"
  ((IS_TERMUX)) && trust_file+="/tls/cert.pem" \
    || trust_file+="/ssl/certs/ca-certificates.crt"
    
  create_file() {
    local file="$1"
    ((INIT)) && rm -f "$file"
    
    if [[ ! -s "$file" ]]; then
      if [[ "$file" =~ gpg$ ]]; then
        echo "$gmail_pass" | "${gpg_cmds[@]}" --symmetric -o "$file"
      else
        local contents=() del='\t' pre=''
        case "$file" in
          *muttrc)
            del='='
            pre='set '
            contents=(
              'sendmail' "$(command -v msmtp)"
              'content_type' 'text/html' 'use_from' 'yes'
              'realname' "$real_name" 'from' "$sender_email"
              'envelope_from' 'yes'
            ) ;;
          *msmtprc)
            contents=(
              'defaults' '' 'auth' 'on' 'tls' 'on'
              'tls_trust_file' "$trust_file" 'account'
              'gmail' 'host' 'smtp.gmail.com' 'port' '587'
              'from' "$sender_email" 'user' "$sender_email"
              'passwordeval' "${gpg_cmds[*]} --decrypt $gpass_file"
              'logfile' "$msmtp_log" 'account default:' 'gmail'
            ) ;;
          *) return 1 ;;
        esac
        printf "${pre}%s${del}%s\n" "${contents[@]}" |
        awk -v d="$del" '{split($0,s,d);if(s[2]==""){print s[1];next}
          if(s[2]~/[[:space:]]/){s[2]="\""s[2]"\""};if(d=="="){print s[1] d s[2];next}
          p=20-length(s[1]);if(p<1)p=1;printf "%s%*s%s\n",s[1],p,"",s[2];next}' > "$file"
      fi
      [[ -f "$file" ]] && chmod 600 "$file"
    fi
  }
  
  eval "$(printf 'create_file "$HOME/.%s"\n' 'gpass.gpg' 'mutt' 'msmtp' |
          sed -E '/\.m/s/"$/rc&/;/(prc|g)"$/{p;s/.*\s//;s|.*/\.(.*)[."]|\1_file=&|
            s/\.gpg_/_/;s/^/local /};/^l.*prc"$/{s/rc_file/_log/;s/rc"$/.log"/}' | sort -r)"
}

send_email() {
  setup_email
  
  local -a recipients copies
  
  if ((DRY_RUN)); then
    recipients=({"yawapisting7","zeenoliev"}"@gmail.com")
    copies=("${recipients[@]}")
  else
    recipients=("jojofundales@"{"hcc.com.ph","yahoo.com"})
    copies=({"arch_rbporral","glachel.arao","rbzden"}"@yahoo.com"
            {"aljonporcalla","eduardo111680"}"@gmail.com")
  fi
  
  local sps=$(spacer 8)
  eval "$(echo '${sps}' |
          sed -E 'p;h;G;G;G;s/\n/ /g;s/ /:/;s/(([^}]+\}){1})([^}]+)/\1\3:5/
            s/(.*) (.*)$/\1\2/;s/^/To/;s|\}$|:2&/g"|;s| |$(sed "s/ /,\\\\n|
            s/$/ <<<"${recipients[*]}")/;p;s/To/CC/;s/recipients/copies/' |
          sed -E '/^[^A-Z]/{s/.$/:5&/;h;G;s/\n/:/;s/^/Subject/;s/$/SSR: $REPORT_DATE/
            p;s/5/2/;s/Subject/Date/;s/SSR.*/$(date +"%B %-d, %Y, %-l:%M %p")/}' |
          sed -E 's/.*/ "&"/;1s/^/details=(\n/;$s/$/\n)/')"
          
  if ((DRY_RUN)) || mutt -s "NSB-P2 SSR as of $REPORT_DATE" \
    -c "$(sed 's/\s/, /g' <<<"${copies[*]}")" \
    -a "$SSR_FILE" -- "${recipients[@]}" \
      < body.html &>/dev/null; then
    send_msg 'Email successfully sent!' "${details[@]}"
  else
    abort 'Failed to send email!' "${details[@]}" \
      '' 'Please check internet connection'
  fi
  
  ((NOT_REPORTED)) && echo "$REPORT_DATE" >> SENT_REPORTS
  return 0
}

# --- Prechecks ---
while (($#)); do
  case "${1:-}" in
    *\.xlsx) SSR_FILE="$1" ;;
    -d|--dry-run) DRY_RUN=1 ;;
    -i|--init) INIT=1 ;;
    *) break ;;
  esac
  shift
done

[[ -s "$SSR_FILE" ]] && SSR_FILE=$(realpath "$SSR_FILE") || \
  abort "Invalid excel file: '$SSR_FILE'" \
    'File does not exists or is corrupted'
[[ "${PREFIX:-}" =~ com.termux ]] || IS_TERMUX=0

# --- Install pkgs ---
install_pkgs

# --- Extract value ---
if ((DRY_RUN)); then
  REPORT_DATE='July 7-13, 2025'
else
  REPORT_DATE=$(xlsx2csv -a --exclude_hidden_sheets "$SSR_FILE" | grep -Eo '^[-]+.*' | tail -n1 | cut -d ' ' -f4- |
                perl -pe 's/\b(\w+)\s+(\d+)-\1\s+(\d+),/$1 $2-$3,/;s/\b(\w)/\u$1/g' | sed -E 's/\s+$//')
fi

grep -q "$REPORT_DATE" SENT_REPORTS && NOT_REPORTED=0
((DRY_RUN+NOT_REPORTED)) && send_email
