const esbuild = require('esbuild');
const path = require('path');

const isProduction = process.argv.includes('--production');
const isWatch = process.argv.includes('--watch');

const config = {
  entryPoints: [path.resolve(__dirname, 'src/extension.ts')],
  bundle: true,
  outfile: path.resolve(__dirname, 'dist/extension.js'),
  platform: 'node',
  format: 'cjs',
  external: ['vscode'],
  sourcemap: !isProduction,
  minify: isProduction,
  target: 'es2020',
};

if (isWatch) {
  esbuild.context(config)
    .then(ctx => ctx.watch())
    .catch(err => {
      console.error(err);
      process.exit(1);
    });
} else {
  esbuild.build(config)
    .catch(err => {
      console.error(err);
      process.exit(1);
    });
}