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

var extractText = function(slideText, slideNumber) {
  var doc = new dom().parseFromString(slideText);
  var ps = xpath.select("//*[local-name()='p']", doc);
  var text = "";
  var sn = slideNumber || null;

  ps.forEach(function(paragraph) {
    paragraph = new dom().parseFromString(paragraph.toString());
    var ts = xpath.select("//*[local-name()='t' or local-name()='tab' or local-name()='br']", paragraph);
    var localText = "";
    ts.forEach(function(t) {
      if (t.localName === "t" && t.childNodes.length > 0 && t.childNodes[0].data != sn) {
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

var getNotes = function(zip, slideFile) {
  var relsText = zip.file(slideFile.replace("ppt/slides", "ppt/slides/_rels") + ".rels").asText();
  var slideNumber = slideFile.replace("ppt/slides/slide", "").replace(".xml", "");
  var doc = new dom().parseFromString(relsText);
  var notesSlide = xpath.select("//*[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide']/@Target", doc)[0];
  if (notesSlide) {
    // return an absolute path within the zipfile rather than a relative path
    var notesPath = notesSlide.value.replace('..', 'ppt');
    var notesText = extractText(zip.file(notesPath).asText(), slideNumber);
    return notesText;
  } else {
    return null;
  }
}

program.args.forEach(function(pptxFile) {
  var cards = [];
  fs.readFile(pptxFile, function(err, data) {
    if (err) throw err;
    var zip = new JSZip(data);
    // console.log(zip);
    Object.keys(zip.files).forEach(function(f) {
      if (slideMatch.test(f)) {
        var slideNumber = f.replace("ppt/slides/slide", "").replace(".xml", "");
        var text = extractText(zip.file(f).asText(), slideNumber);
        var notesText = getNotes(zip, f);
        cards.push({
          slideNumber: slideNumber,
          text: text,
          notes: notesText
        });
      }
    });
    cards.sort(function(a, b) {
      return a.slideNumber - b.slideNumber
    });
    jsonfile.writeFileSync(pptxFile + '.json', cards, {
      spaces: 2
    });
  });
})
