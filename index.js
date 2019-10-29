#!/usr/bin/env node
const Admzip = require('adm-zip')
const path = require('path')
const program = require('commander')
const chalk = require('chalk')
const fs = require('fs')
const shelljs = require('shelljs')

const { exec } = shelljs

const HEADTMPLATE = `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="X-UA-Compatible" content="ie=edge">
  <title>用户服务协议</title>
  <style>
    * {
      margin: 0;
      padding: 0;
    }
    body {
      font-size: 14px;
    }
    p {
      font-size: 0.14rem;
      color: rgba(0,0,0,.8);
    }
    .container {
      padding: 0.2rem 0.2rem 0.8rem 0.2rem;
    }
    .title {
      text-align: center;
      line-height: 1.8;
      font-size: 0.25rem;
      color: #000;
      font-weight: 400;
    }
    .title-2 {
      font-weight: bold;
    }
    /* .p::first-line {
      text-indent: 0.5rem;
    } */
    a {
      color: #208dff;
    }
  </style>
</head>
<body>
  <div class="container">
`
const FOOTERTMPLATE = `
  </div>
<script>
    function initRem() {
      var clientWidth = document.documentElement.clientWidth
      var rem = clientWidth / 750
      document.documentElement.style.fontSize = rem * 100 + 'px'
    }
    function addIndent() {
      var paragraphs = document.getElementsByClassName('p')
      for(var i = 0; i < paragraphs.length; i++) {
        var textNode = paragraphs[i].firstChild
        var span = document.createElement('span')
        span.innerText = 'span'
        span.style.color = '#fff'
        span.style.display = 'inline-block'
        span.style.width = '0.3rem'
        paragraphs[i].insertBefore(span, textNode)
      }
    }
    initRem()
    addIndent()
    window.addEventListener('resize', initRem, false)
</script>
</body>
</html>
`
// resolve argvs
program.version('1.0.0')
.option('-f, --file <value>', 'filename')
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