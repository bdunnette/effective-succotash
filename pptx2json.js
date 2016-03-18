#!/usr/bin/env node

var xpath = require('xpath'),
  dom = require('xmldom').DOMParser,
  JSZip = require("jszip"),
  fs = require('fs'),
  path = require('path'),
  program = require('commander'),
  jsonfile = require('jsonfile'),
  slideMatch = /^ppt\/slides\/slide/;

program.parse(process.argv);

var extractText = function(slideText) {
  var doc = new dom().parseFromString(slideText);
  var ps = xpath.select("//*[local-name()='p']", doc);
  var text = "";

  ps.forEach(function(paragraph) {
    paragraph = new dom().parseFromString(paragraph.toString());
    var ts = xpath.select("//*[local-name()='t' or local-name()='tab' or local-name()='br']", paragraph);
    var localText = "";
    ts.forEach(function(t) {
      if (t.localName === "t" && t.childNodes.length > 0) {
        localText += t.childNodes[0].data;
      } else {
        if (t.localName === "tab" || t.localName === "br") {
          localText += "";
        }
      }
    });
    text += localText + "\n";
  });

  return text;

};

program.args.forEach(function(pptxFile) {
  var cards = [];
  console.log(pptxFile);
  fs.readFile(pptxFile, function(err, data) {
    if (err) throw err;
    var zip = new JSZip(data);
    // console.log(zip);
    Object.keys(zip.files).forEach(function(f) {
      if (slideMatch.test(f)) {
        console.log(f);
        var slide = f.replace("ppt/slides/slide", "").replace(".xml", "");
        var slideText = zip.file(f).asText();
        var text = extractText(slideText);
        // console.log(slideText);
        cards.push({
          slide: slide,
          text: text,
          fullText: slideText
        });
      }
    });
    // console.log(cards);
    jsonfile.writeFileSync(pptxFile + '.json', cards, {
      spaces: 2
    });
  });
})
