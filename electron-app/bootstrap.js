const path = require('path')
const fs = require('fs')
const dotenv = require('dotenv')

const parsedEnv = dotenv.parse(fs.readFileSync(path.join(__dirname, '.env')))
Object.assign(process.env, parsedEnv)