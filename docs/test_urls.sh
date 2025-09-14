#!/bin/bash
echo "Running URL tests for Tarun's Omexom project..."
echo ""

BASE_URL="https://lfa-lab.github.io/Omexom/"

echo "------------------------------------------"
echo "1. Testing Importsimple.bas"
echo "------------------------------------------"
curl -I "${BASE_URL}github-pages/Importsimple.bas"
if [ $? -ne 0 ]; then
  echo "ERROR: Failed to access the file."
else
  echo "SUCCESS: File is accessible."
fi
echo ""

echo "------------------------------------------"
echo "2. Testing TemplateProject_v1.mpt"
echo "------------------------------------------"
curl -I "${BASE_URL}github-pages/TemplateProject_v1.mpt"
if [ $? -ne 0 ]; then
  echo "ERROR: Failed to access the file."
else
  echo "SUCCESS: File is accessible."
fi
echo ""

echo "------------------------------------------"
echo "3. Testing index.json"
echo "------------------------------------------"
curl -I "${BASE_URL}github-pages/index.json"
if [ $? -ne 0 ]; then
  echo "ERROR: Failed to access the file."
else
  echo "SUCCESS: File is accessible."
fi
echo ""

echo "------------------------------------------"
echo "4. Testing documentation in docs/ folder (e.g., index.html)"
echo "------------------------------------------"
curl -I "${BASE_URL}docs/index.html"
if [ $? -ne 0 ]; then
  echo "ERROR: Failed to access the file."
else
  echo "SUCCESS: File is accessible."
fi
echo ""

echo "Test complete. If you see 'HTTP/2 200' for each file, the test passed."