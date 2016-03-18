#!/usr/bin/env node

var xpath = require('xpath'),
  dom = require('xmldom').DOMParser,
  JSZip = require("jszip"),
  fs = require('fs'),
  path = require('path'),
  program = require('commander'),
  slideMatch = /^ppt\/slides\/slide/;

program.parse(process.argv);

console.log(program.args);

program.args.forEach(function(pptxFile) {
  var cards = [];
  console.log(pptxFile);
  fs.readFile(pptxFile, function(err, data) {
    if (err) throw err;
    var zip = new JSZip(data);
    console.log(zip);
    Object.keys(zip.files).forEach(function(f) {
      if (slideMatch.test(f)) {
        console.log(f);
        var slideText = zip.file(f).asText();
        console.log(slideText);
        cards.push(slideText);
      }
    });
    console.log(cards);
  });
})
