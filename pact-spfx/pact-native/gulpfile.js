'use strict';

const { spawn } = require('node:child_process');
const path = require('node:path');
const gulp = require('gulp');

const projectRoot = __dirname;
const heftBin = path.join(projectRoot, 'node_modules', '@rushstack', 'heft', 'bin', 'heft');

function runHeft(args) {
  return (done) => {
    const child = spawn(process.execPath, [heftBin, ...args], {
      cwd: projectRoot,
      stdio: 'inherit',
      shell: false
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
