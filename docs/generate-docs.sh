#!/bin/bash

# Generate PDF documentation for OpenClaw MS Teams User Plugin
# Requires: sudo gem install asciidoctor-pdf

echo "OpenClaw MS Teams User Plugin — Documentation Generator"
echo "======================================================="
echo ""

# Check if asciidoctor-pdf is installed
if ! command -v asciidoctor-pdf &> /dev/null; then
    echo "Error: asciidoctor-pdf is not installed"
    echo "Please install it with: sudo gem install asciidoctor-pdf"
    exit 1
fi

# Change to docs directory
cd "$(dirname "$0")"

THEME_ARGS="-a pdf-theme=knowall -a pdf-themesdir=themes -a pdf-fontsdir=themes"
ERRORS=0

for doc in SOLUTION_DESIGN DEPLOYMENT TESTING TROUBLESHOOTING; do
    if [ -f "${doc}.adoc" ]; then
        echo "Generating ${doc}.pdf..."
        asciidoctor-pdf $THEME_ARGS "${doc}.adoc" -o "${doc}.pdf"
        if [ $? -eq 0 ]; then
            echo "  Done: ${doc}.pdf"
        else
            echo "  FAILED: ${doc}.adoc"
            ERRORS=$((ERRORS + 1))
        fi
    fi
done

echo ""
if [ $ERRORS -eq 0 ]; then
    echo "All documents generated successfully!"
else
    echo "${ERRORS} document(s) failed to generate."
    exit 1
fi
