// index.src.html 의 <script src="core.js"> 를 core.js 내용으로 바꿔 단일 파일 index.html 생성
const fs = require('fs');
const core = fs.readFileSync(__dirname + '/core.js', 'utf8');
if (core.includes('</script')) throw new Error('core.js 에 </script 가 들어 있어요');
const src = fs.readFileSync(__dirname + '/index.src.html', 'utf8');
const tag = '<script src="core.js"></script>';
if (!src.includes(tag)) throw new Error('script 태그를 찾을 수 없어요');
fs.writeFileSync(__dirname + '/index.html', src.replace(tag, () => '<script>\n' + core + '\n</script>'));
console.log('index.html 생성 완료 (단일 파일)');
