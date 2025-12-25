const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const packageJsonPath = path.join(__dirname, '..', 'package.json');

try {
  // 执行 build（package.json 会自动作为资源文件被包含）
  console.log('开始构建...');
  execSync('ng build', { stdio: 'inherit' });

  // build 成功后才更新版本号
  console.log('\n构建成功，更新版本号...');

  // 重新读取 package.json（构建过程中可能被修改）
  const updatedPackageJson = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));

  // 解析版本号
  const versionParts = updatedPackageJson.version.split('.');
  const lastPart = parseInt(versionParts[versionParts.length - 1], 10);

  // 将最后一位 +1
  versionParts[versionParts.length - 1] = (lastPart + 1).toString();

  // 更新版本号
  updatedPackageJson.version = versionParts.join('.');

  // 写回 package.json
  fs.writeFileSync(packageJsonPath, JSON.stringify(updatedPackageJson, null, 2) + '\n', 'utf8');

  console.log(`✓ 版本号已更新: ${updatedPackageJson.version}`);
} catch (error) {
  console.error('\n构建失败，版本号未更新');
  process.exit(1);
}

