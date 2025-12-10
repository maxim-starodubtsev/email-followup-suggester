#!/bin/bash

# Script to trust the webpack dev server SSL certificate in macOS Keychain
# This fixes the "invalid security certificate" error in Outlook

CERT_PATH="node_modules/.cache/webpack-dev-server/server.pem"
CERT_NAME="webpack-dev-server"

echo "🔐 Trusting webpack dev server certificate for Outlook..."
echo ""

# Check if certificate exists
if [ ! -f "$CERT_PATH" ]; then
    echo "❌ Certificate not found at $CERT_PATH"
    echo "   Please start the dev server first (npm start) to generate the certificate."
    exit 1
fi

# Extract certificate from PEM file
TEMP_CERT="/tmp/${CERT_NAME}.cer"
openssl x509 -in "$CERT_PATH" -outform DER -out "$TEMP_CERT" 2>/dev/null

if [ $? -ne 0 ]; then
    echo "❌ Failed to extract certificate"
    exit 1
fi

# Remove existing certificate from Keychain if it exists
security delete-certificate -c "$CERT_NAME" login.keychain 2>/dev/null

# Add certificate to login keychain
security add-certificates "$TEMP_CERT" 2>/dev/null

if [ $? -ne 0 ]; then
    echo "❌ Failed to add certificate to Keychain"
    rm -f "$TEMP_CERT"
    exit 1
fi

# Find the certificate and set it to "Always Trust"
CERT_SHA=$(security find-certificate -c "$CERT_NAME" -a login.keychain 2>/dev/null | grep "SHA-1" | head -1 | awk '{print $3}')

if [ -n "$CERT_SHA" ]; then
    # Set trust settings to "Always Trust"
    security add-trusted-cert -d -r trustRoot -k login.keychain "$TEMP_CERT" 2>/dev/null
    
    if [ $? -eq 0 ]; then
        echo "✅ Certificate added to Keychain and set to 'Always Trust'"
        echo ""
        echo "📝 Next steps:"
        echo "   1. Restart Outlook"
        echo "   2. Try opening the add-in again"
        echo ""
        echo "💡 If you still see certificate errors:"
        echo "   1. Open Keychain Access app"
        echo "   2. Search for 'webpack-dev-server'"
        echo "   3. Double-click the certificate"
        echo "   4. Expand 'Trust' section"
        echo "   5. Set 'When using this certificate' to 'Always Trust'"
    else
        echo "⚠️  Certificate added but trust settings may need manual configuration"
        echo "   Please open Keychain Access and set the certificate to 'Always Trust'"
    fi
else
    echo "⚠️  Certificate added but may need manual trust configuration"
    echo "   Please open Keychain Access and set the certificate to 'Always Trust'"
fi

# Clean up
rm -f "$TEMP_CERT"

echo ""
echo "🔍 To verify, open Keychain Access and search for 'webpack-dev-server'"







