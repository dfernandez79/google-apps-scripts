// Replace this variable with the folder used to get the images:
//
// Option 1: IMAGES_FOLDER = {id: 'GOOGLE DRIVE FOLDER ID'}
// Option 2: IMAGES_FOLDER = {name: 'Sibling Folder Name' }
//
// NOTE: The variable is not set directly because when the doc is opening, executing
//       some DriveApp methods could fail due to auth. That failure makes the
//       onOpen function to not be run.
//
var IMAGES_FOLDER = {name: 'Mockups' };
var IGNORE_PATTERN = /^-/;
var MAX_WIDTH = 677;
var FILE_ID_REGEX = /https:\/\/drive.google.com\/.*\/file\/d\/(.*)\/.*/;

function onOpen (e) {
  DocumentApp.getUi()
    .createAddonMenu()
    .addItem('Synchronize', 'confirmImageSynchronization')
    .addItem('Update selected image', 'updateSelected')
    .addToUi();
}

function findFolder (name) {
  var doc = DocumentApp.getActiveDocument();
  var file = DriveApp.getFileById(doc.getId());
  var parent = file.getParents().next();
  return parent.getFoldersByName(name).next();
}

function updateSelected () {
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    var elements = selection.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
      updateElement(elements[i].getElement().asInlineImage());
    }
  }
}

function updateElement (image) {
  if (image) {
    var match = FILE_ID_REGEX.exec(image.getLinkUrl());
    if (match) {
      var file = DriveApp.getFileById(match[1]);
      replaceImage(image, file.getBlob(), true, file.getName().endsWith("@2x.png"), MAX_WIDTH);
    }
  }
}

function replaceImage (image, blob, updateSize, hidpi, maxWidth) {
  var childIndex = image.getParent().getChildIndex(image);
  var newImg = image.getParent().insertInlineImage(childIndex, blob);

  if (!updateSize) {
    newImg.setWidth(image.getWidth());
    newImg.setHeight(image.getHeight());
  } else {
    var width = hidpi ? newImg.getWidth() / 2 : newImg.getWidth();
    var height = hidpi ? newImg.getHeight() / 2 : newImg.getHeight();

    if (width > maxWidth) {
      ratio = height / width;
      newImg.setWidth(maxWidth);
      newImg.setHeight(Math.floor(maxWidth * ratio));
    } else {
      newImg.setWidth(width);
      newImg.setHeight(height);
    }
  }

  newImg.setLinkUrl(image.getLinkUrl());
  image.removeFromParent();
}

function confirmImageSynchronization () {
  var ui = DocumentApp.getUi();
  var folder;

  if ('id' in IMAGES_FOLDER) {
    folder = DriveApp.getFolderById(IMAGES_FOLDER.id);
  } else if ('name' in IMAGES_FOLDER) {
    folder = findFolder(IMAGES_FOLDER.name);
  } else {
    throw new Error('The IMAGES_FOLDER variable is not correctly configured');
  }

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
  var updated = {};
  var imagesToSync = this.body.getImages().filter(function (img) { return imagesByLink[img.getLinkUrl()]; });

  imagesToSync.forEach(function (img) {
    replaceImage(img, imagesByLink[img.getLinkUrl()].getBlob());
    updated[img.getLinkUrl()] = true;
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