const assert = require('node:assert')
const fs = require('node:fs')
const path = require('node:path')
const os = require('node:os')
const { fork } = require('node:child_process')

const projectRoot = path.resolve(__dirname, '..', '..')
const workerPath = path.join(projectRoot, 'electron', 'import-worker.cjs')

function runWorker(filePath, userData) {
  return new Promise((resolve, reject) => {
    const worker = fork(workerPath, [], {
      env: {
        ...process.env,
        ELECTRON_RUN_AS_NODE: '1',
        ERGONOMIC_USER_DATA: userData,
      },
      stdio: ['ignore', 'pipe', 'pipe', 'ipc'],
    })

    let stderr = ''

    worker.stderr?.on('data', (d) => {
      stderr += String(d)
    })

    const timeout = setTimeout(() => {
      worker.kill()
      reject(new Error('Worker timed out'))
    }, 30000)

    worker.once('message', (msg) => {
      clearTimeout(timeout)
      resolve({ msg, stderr })
    })

    worker.once('error', reject)

    worker.send({ filePath, userData })
  })
}

async function main() {
  console.log('Starting import-worker smoke test...\n')

  const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'erg-import-test-'))
  const userData = path.join(tempRoot, 'userData')

  fs.mkdirSync(userData, { recursive: true })

  // TXT TEST
  const txtPath = path.join(tempRoot, 'sample.txt')
  fs.writeFileSync(txtPath, 'Hello ergonomic import test.', 'utf8')

  const txtResult = await runWorker(txtPath, userData)

  console.log('TXT RESULT:')
  console.log(JSON.stringify(txtResult.msg, null, 2))

  assert.equal(txtResult.msg.ok, true)

  // PNG TEST
  const pngPath = path.join(tempRoot, 'sample.png')

  const onePixelPng = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=',
    'base64'
  )

  fs.writeFileSync(pngPath, onePixelPng)

  const pngResult = await runWorker(pngPath, userData)

  console.log('\nPNG RESULT:')
  console.log(JSON.stringify(pngResult.msg, null, 2))

  assert.equal(pngResult.msg.ok, true)

  // BROKEN PDF TEST
  const badPdfPath = path.join(tempRoot, 'broken.pdf')

  fs.writeFileSync(
    badPdfPath,
    Buffer.from('%PDF-1.4\nbroken pdf\n%%EOF')
  )

  const badPdfResult = await runWorker(badPdfPath, userData)

  console.log('\nBROKEN PDF RESULT:')
  console.log(JSON.stringify(badPdfResult.msg, null, 2))

  assert.ok(
  badPdfResult.msg.ok === true || badPdfResult.msg.error,
  `Expected safe success or safe error, got: ${JSON.stringify(badPdfResult.msg, null, 2)}`
)

  console.log('\n✅ import-worker smoke test passed')
}

main().catch((error) => {
  console.error('\n❌ import-worker smoke test failed')
  console.error(error)
  process.exit(1)
})