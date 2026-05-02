'use strict';

const { spawn } = require('node:child_process');
const path = require('node:path');
const gulp = require('gulp');

const projectRoot = __dirname;
const heftBin = process.platform === 'win32'
  ? path.join(projectRoot, 'node_modules', '.bin', 'heft.cmd')
  : path.join(projectRoot, 'node_modules', '.bin', 'heft');

function runHeft(args) {
  return (done) => {
    // On Windows, .cmd files need shell:true. Wrap bin path in quotes to handle spaces.
    const quotedBin = process.platform === 'win32' ? `"${heftBin}"` : heftBin;
    const child = spawn(quotedBin, args, {
      cwd: projectRoot,
      stdio: 'inherit',
      shell: true,
      windowsVerbatimArguments: false
    });

    child.on('error', done);
    child.on('exit', (code) => {
      if (code === 0) {
        done();
        return;
      }

      done(new Error(`Heft exited with code ${code}`));
    });
  };
}

gulp.task('serve', runHeft(['start', '--clean']));
gulp.task('build', runHeft(['test', '--clean', '--production']));
gulp.task('clean', runHeft(['clean']));
gulp.task('package-solution', runHeft(['package-solution', '--production']));
gulp.task('test', runHeft(['test', '--clean']));
gulp.task('default', gulp.series('serve'));
