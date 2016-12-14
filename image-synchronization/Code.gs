// Note replace this variable with the folder used to get the images:
//
// Option 1: Using a folder ID
//   DriveApp.getFolderById('FOLDER_ID');
//
// Option 2: Using the name of a folder contained in this doc parent
// (note this method may fail if the doc is linked to multiple parents)
//
//   findFolder('FOLDER NAME')
//
var IMAGES_FOLDER = findFolder('Mockups');
var IGNORE_PATTERN = /^-/;
var MAX_WIDTH = 677;

function onOpen (e) {
  DocumentApp.getUi()
    .createAddonMenu()
    .addItem('Synchronize', 'confirmImageSynchronization')
    .addToUi();
}

function findFolder (name) {
  var doc = DocumentApp.getActiveDocument();
  var file = DriveApp.getFileById(doc.getId());
  var parent = file.getParents().next();
  return parent.getFoldersByName(name).next();
}

function confirmImageSynchronization () {
  var ui = DocumentApp.getUi();
  var folder = IMAGES_FOLDER;

  if (ui.alert(
      'Image Synchronization',
      'Update this document with images in "' + folder.getName() + '"?', ui.ButtonSet.YES_NO) === ui.Button.YES) {

    new ImageSynchronization(folder, DocumentApp.getActiveDocument(), IGNORE_PATTERN, MAX_WIDTH).execute();
  }
}

function ImageSynchronization (folder, document, ignorePattern, maxWidth) {
  this.folder = folder;
  this.body = document.getBody();
  this.ignorePattern = ignorePattern;
  this.maxWidth = maxWidth;
}

ImageSynchronization.prototype.execute = function () {
  var imagesByLink = this._collectImages();

  var updated = this._updateImages(imagesByLink);
  this._appendImages(imagesByLink, updated);
}

ImageSynchronization.prototype._collectImages = function () {
  var files = this.folder.getFilesByType('image/png');
  var result = {};
  var file = null;
  while (files.hasNext()) {
    file = files.next();
    if (!this.ignorePattern.test(file.getName())) {
      result[file.getUrl()] = file;
    }
  }
  return result;
}

ImageSynchronization.prototype._updateImages = function (imagesByLink) {
  var updated = {}, childIndex;
  var imagesToSync = this.body.getImages().filter(function (img) { return imagesByLink[img.getLinkUrl()]; });

  imagesToSync.forEach(function (img) {
    childIndex = img.getParent().getChildIndex(img);
    var newImg = img.getParent().insertInlineImage(childIndex, imagesByLink[img.getLinkUrl()].getBlob());
    newImg.setWidth(img.getWidth());
    newImg.setHeight(img.getHeight());
    newImg.setLinkUrl(img.getLinkUrl());
    updated[img.getLinkUrl()] = true;
    img.removeFromParent();
  });

  return updated;
}

ImageSynchronization.prototype._appendImages = function (imagesByLink, ignore) {
  var img, ratio;
  var self = this;

  this._sortImagesByFileName(imagesByLink).forEach(function (file) {
    if (!ignore[file.getUrl()]) {
      img = self.body.appendImage(file.getBlob());
      img.setLinkUrl(file.getUrl());

      var hidpi = file.getName().endsWith('@2x.png');
      var width = hidpi ? img.getWidth() / 2 : img.getWidth();
      var height = hidpi ? img.getHeight() / 2 : img.getHeight();

      if (width > self.maxWidth) {
        ratio = height / width;
        img.setWidth(self.maxWidth);
        img.setHeight(Math.floor(self.maxWidth * ratio));
      } else {
        img.setWidth(width);
        img.setHeight(height);
      }
    }
  });
}

ImageSynchronization.prototype._sortImagesByFileName = function (imagesByLink) {
  return Object.keys(imagesByLink).sort(function (linkA, linkB) {
    var fileNameA = imagesByLink[linkA].getName();
    var fileNameB = imagesByLink[linkB].getName();
    if (fileNameA > fileNameB) {
      return 1;
    }
    if (fileNameA < fileNameB) {
      return -1;
    }
    return 0;
  }).map(function (key) {
    return imagesByLink[key];
  });
}

// String.endsWith polyfill
// See: https://developer.mozilla.org/en/docs/Web/JavaScript/Reference/Global_Objects/String/endsWith
if (!String.prototype.endsWith) {
  String.prototype.endsWith = function(searchString, position) {
      var subjectString = this.toString();
      if (typeof position !== 'number' || !isFinite(position) || Math.floor(position) !== position || position > subjectString.length) {
        position = subjectString.length;
      }
      position -= searchString.length;
      var lastIndex = subjectString.lastIndexOf(searchString, position);
      return lastIndex !== -1 && lastIndex === position;
  };
}