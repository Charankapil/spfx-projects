#!/usr/bin/env bash
# Makes the sign-in certificate for the access report app registration.
#
#   ./new-certificate.sh [name] [days]
#
# Writes into the current folder:
#   <name>.cer          public certificate  -> upload to the app registration (Certificates & secrets)
#   <name>.secret.pem   certificate + private key -> the value of ACCESS_REPORT_CERTIFICATE
#
# Keep <name>.secret.pem out of source control and delete it once it is stored as a secret.
set -euo pipefail

name="${1:-access-report}"
days="${2:-730}"

if ! command -v openssl >/dev/null 2>&1; then
  echo "openssl is required." >&2
  exit 1
fi
if [ -e "$name.secret.pem" ] || [ -e "$name.cer" ]; then
  echo "$name.secret.pem or $name.cer already exists; choose another name or remove them first." >&2
  exit 1
fi

umask 077
tmp="$(mktemp -d)"
trap 'rm -rf "$tmp"' EXIT

openssl req -x509 -newkey rsa:2048 -sha256 -nodes \
  -keyout "$tmp/key.pem" -out "$tmp/cert.pem" \
  -days "$days" -subj "/CN=$name" >/dev/null 2>&1

cat "$tmp/cert.pem" "$tmp/key.pem" > "$name.secret.pem"
cp "$tmp/cert.pem" "$name.cer"
chmod 644 "$name.cer"

echo "Created:"
echo "  $name.cer         upload this to the app registration"
echo "  $name.secret.pem  store the whole file content as the secret ACCESS_REPORT_CERTIFICATE"
echo
echo "Thumbprint (as the Entra admin center shows it after upload):"
openssl x509 -in "$tmp/cert.pem" -noout -fingerprint -sha1 | sed 's/.*=//; s/://g'
echo "Expires:"
openssl x509 -in "$tmp/cert.pem" -noout -enddate | sed 's/notAfter=//'
