#!/usr/bin/env node
const Admzip = require('adm-zip')
const path = require('path')
const program = require('commander')
const chalk = require('chalk')
const fs = require('fs')
const shelljs = require('shelljs')

const { exec } = shelljs

const HEADTMPLATE = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="X-UA-Compatible" content="ie=edge">
  <title>title</title>
</head>
<body>
  <div class="container">
`
const FOOTERTMPLATE = `
  </div>
</body>
</html>
`
// resolve argvs
program.version('1.0.0')
.option('-f, --filename <value>', 'filename')
.parse(process.argv)

const filename = program.filename

if (filename) {
  convertSingle(filename)
} else {
  convertAll()
}

function convertSingle(filename) {
  exec('rm -rf ./tmp')
  fs.mkdirSync('./tmp')
  const zip = new Admzip(getCurrPath(filename))
    zip.extractAllTo(getCurrPath('./tmp/'), true)
  const documentText = fs.readFileSync(getCurrPath('./tmp/word/document.xml'), {
    encoding: 'utf-8'
  })
  const paragraphs = documentText.match(/<w:p\s.*?<\/w:p>/g).filter(p => p)
  const results = []
  for(let i = 0, l = paragraphs.length; i < l; i++) {
    const texts = paragraphs[i].match(/<w:t>.*?<\/w:t>/g)
    if (Array.isArray(texts) && texts.length) {
      const paragraphText = texts.map(t => t.slice(5, -6)).join('')
      const pText = `    <p class="item">${paragraphText}</p>\n`
      results.push(pText)
    }
  }
  results[results.length - 1] = results[results.length - 1].replace('\n', '')
  const output = HEADTMPLATE + results.join('') + FOOTERTMPLATE
  exec('mkdir wordConvert')
  fs.writeFileSync(`./wordConvert/${filename}.html`, output, {
    encoding: 'utf-8'
  })
  exec('rm -rf tmp')
}

function convertAll() {
  const files = fs.readdirSync(process.cwd())
  files.filter((filename) => /\.docx$/.test(filename)).forEach(filename => convertSingle(filename))
}

function getCurrPath(p) {
  return path.resolve(process.cwd(), p)
}