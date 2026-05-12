const { spawn } = require('node:child_process')
const electron = require('electron')

const env = { ...process.env }
delete env.ELECTRON_RUN_AS_NODE

const args = process.argv.slice(2).filter((arg) => {
  if (arg.startsWith('--dev-server=')) {
    env.ERGONOMIC_DEV_SERVER_URL = arg.slice('--dev-server='.length)
    return false
  }
  return true
})

const child = spawn(electron, args, {
  stdio: 'inherit',
  env,
})

child.on('exit', (code, signal) => {
  if (signal) process.kill(process.pid, signal)
  process.exit(code ?? 0)
})
