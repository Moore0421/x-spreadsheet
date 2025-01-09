const chokidar = require('chokidar');
const spawn = require('cross-spawn');
const path = require('path');

let building = false;
let pendingBuild = false;
let currentProcess = null;

// 定义要监听的文件
const watchPaths = [
  'src/**/*.js',
  'src/**/*.vue',
  'src/**/*.jsx',
  'src/**/*.ts',
  'src/**/*.tsx',
  'index.html',
  "server.js",
  "upload.py",
  "package.json",
  "data.js"
];

// 排除的文件/文件夹
const ignoredPaths = [
  'node_modules',
  'dist',
  '.git',
  'build'
];

function runBuild() {
  if (building) {
    pendingBuild = true;
    return;
  }

  building = true;
  console.log('\n🚀 开始构建...');

  // 如果有正在运行的进程，先终止它
  if (currentProcess) {
    currentProcess.kill();
  }

  // 运行 webpack 构建
  currentProcess = spawn('npm', ['run', 'build'], {
    stdio: 'inherit',
    shell: true
  });

  currentProcess.on('close', (code) => {
    building = false;
    currentProcess = null;

    if (code === 0) {
      console.log('✅ 构建完成');
    } else {
      console.log('❌ 构建失败');
    }

    // 如果在构建过程中有新的变更，立即开始新的构建
    if (pendingBuild) {
      pendingBuild = false;
      runBuild();
    }
  });
}

// 创建文件监听器
const watcher = chokidar.watch(watchPaths, {
  ignored: ignoredPaths,
  persistent: true,
  ignoreInitial: true,
  awaitWriteFinish: {
    stabilityThreshold: 100,
    pollInterval: 100
  }
});

// 监听文件变化事件
watcher
  .on('change', (filePath) => {
    const relativePath = path.relative(process.cwd(), filePath);
    console.log(`\n📝 文件变更: ${relativePath}`);
    runBuild();
  })
  .on('ready', () => {
    console.log('👀 开始监听文件变更...');
    // 首次运行构建
    runBuild();
  });

// 处理进程退出
process.on('SIGINT', () => {
  if (currentProcess) {
    currentProcess.kill();
  }
  watcher.close();
  process.exit();
}); 