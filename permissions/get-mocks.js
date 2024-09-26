#! /usr/bin/env node

import fs from 'fs';
import path from 'path';

const commandsPath = '../dist';

// stub describe to avoid errors when loading the test file
global.describe = () => { };

function getAllTestFiles(dirPath, arrayOfFiles = []) {
  const files = fs.readdirSync(dirPath);

  files.forEach((file) => {
    const filePath = path.join(dirPath, file);
    if (fs.statSync(filePath).isDirectory()) {
      arrayOfFiles = getAllTestFiles(filePath, arrayOfFiles);
    }
    else if (file.endsWith('.spec.js')) {
      arrayOfFiles.push(filePath);
    }
  });

  return arrayOfFiles;
}

function encodeQueryStringValue(value) {
  return value
    .replace(/ /g, '%20')
    .replace(/'/g, '%27')
    .replace(/@/g, '%40');
}

const mocks = [];
const jsFiles = getAllTestFiles(commandsPath);
for (const file of jsFiles) {
  console.log(`Processing ${file}`);
  try {
    const mocksFromFile = (await import(file)).mocks;
    if (!mocksFromFile) {
      continue;
    }
    const mockValues = Object.values(mocksFromFile);
    mockValues.forEach(mock => {
      mock.src = path.relative(commandsPath, file);
    });
    mocks.push(...mockValues);
  }
  catch (e) {
    console.error(`Error processing ${file}: ${e}`);
  }
}

// encode URLs
mocks.forEach(mock => {
  const [baseUrl, queryString] = mock.request.url.split('?');
  if (!queryString) {
    return;
  }

  const params = new URLSearchParams(queryString);
  const encodedQueryString = Array.from(params.entries())
    .map(value => `${value[0]}=${encodeQueryStringValue(value[1])}`)
    .join('&');

  mock.request.url = `${baseUrl}?${encodedQueryString}`;
});

// group mocks by the host header in the request url
const groupedMocks = mocks.reduce((acc, mock) => {
  const host = new URL(mock.request.url).host;
  if (!acc[host]) {
    acc[host] = [];
  }
  acc[host].push(mock);
  return acc;
}, {});

for (const host in groupedMocks) {
  // sort mocks by the length of the request url so that more specific mocks are at the top
  groupedMocks[host].sort((a, b) => b.request.url.length - a.request.url.length);
  // todo: possibly save long bodies to separate files
  const mocksFileContents = {
    "$schema": "https://raw.githubusercontent.com/microsoft/dev-proxy/main/schemas/v0.20.1/mockresponseplugin.schema.json",
    "mocks": groupedMocks[host]
  };
  fs.writeFileSync(path.join('.devproxy', `mocks-${host}.json`), JSON.stringify(mocksFileContents, null, 2));
}