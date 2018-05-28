'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');
const spsync = require('gulp-spsync-creds').sync;

// const environmentInfo = {
//     "username": "thyt@sorab365.onmicrosoft.com",
//     "password": "Linn2004",
//     "tenant": "Sorab365",
//     "catalogSite": "/sites/appcatalog"
// }
// const environmentInfo = {
//     "username": "Consid_Adm_2@elofhanssonab.onmicrosoft.com",
//     "password": "Linn2004",
//     "tenant": "elofhanssonab",
//     "catalogSite": "/sites/appcatalog"
// }
const environmentInfo = {
    "username": "conthyt@returpack.se",
    "password": "Linn20040204",
    "tenant": "returpack",
    "catalogSite": "/sites/ledningsystem"
}

build.task('upload-app-pkg', {
    execute: (config) => {
      environmentInfo.username = config.args['username'] || environmentInfo.username;
      environmentInfo.password = config.args['password'] || environmentInfo.password;
      environmentInfo.tenant = config.args['tenant'] || environmentInfo.tenant;
      environmentInfo.catalogSite = config.args['catalogsite'] || environmentInfo.catalogSite;

      return new Promise((resolve, reject) => {
        const pkgFile = require('./config/package-solution.json');
        const folderLocation = `./sharepoint/${pkgFile.paths.zippedPackage}`;

        return gulp.src(folderLocation)
          .pipe(spsync({
            "username": environmentInfo.username,
            "password": environmentInfo.password,
            "site": `https://${environmentInfo.tenant}.sharepoint.com/${environmentInfo.catalogSite}`,
            "libraryPath": "AppCatalog",
            "publish": true
          }))
          .on('finish', resolve);
      });
    }
  });
