import request from 'request'
import xlsx from 'node-xlsx'
import path from 'path'
import untildify from 'untildify'

const { apiKey, path: filePath, wid = 605711, pid = 91353874 } = require('optimist').argv;

let entries = []

const getRequestData = ({ issueKey, workDescription, workDate, hours }) => {
  const startDate = new Date(workDate)
  const duration = parseInt(hours, 10) * 60 * 60

  return {
    method: 'POST',
    url: 'https://www.toggl.com/api/v8/time_entries',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Basic ${Buffer.from(`${apiKey}:api_token`).toString('base64')}`,
    },
    body: JSON.stringify({
      time_entry: {
        billable: true,
        created_with: 'Snowball',
        description: `${issueKey} - ${workDescription}`,
        duration,
        pid,
        start: startDate.toISOString(),
        stop: new Date(startDate.getTime() + 1000 * duration).toISOString(),
        tid: null,
        wid,
      }
    }),
  }
}

function addToArray(entry) {
  entries = [...entries, entry]
}

function processEntriesToToggl() {
  entries.forEach((entry, index) => {
    setTimeout(() => request(getRequestData(entry), (error) => {
      if (error) throw new Error(error);
      else console.log(`Done entry ${entry.issueKey} from ${entry.workDate.toISOString()}`)
    }), 1500 * index)
  })
}

function convertExcelDate(date) {
  return new Date(Math.round((date - (25567 + 2))*86400*1000))
}

function readFromExcel() {
  try {
    const doc = xlsx.parse(path.normalize(untildify(filePath)));
    const worklogs = doc.filter(sheet => sheet.name === 'Worklogs')

    if (worklogs.length === 0) {
      console.log('The sheet with hours must be called Worklogs')

      return
    }

    const columnsToPick = ['Issue Key', 'Hours', 'Work date', 'Work Description']
    const columnsObjectNames = ['issueKey', 'hours', 'workDate', 'workDescription']
    const columnsIndices = []
    const dateIndices = [2]

    const parseColumnHeaders = data => {
      data.forEach((column, index) => ~columnsToPick.indexOf(column) && columnsIndices.push(index))
    }

    worklogs[0].data.forEach((data, index) => {
      if (index === 0) {
        parseColumnHeaders(data)

        return
      }

      addToArray(columnsIndices.reduce((prev, cur, index) => ({
        ...prev,
        [columnsObjectNames[index]]: ~dateIndices.indexOf(index) ? convertExcelDate(data[cur]) : data[cur]
      }), {}))
    })

    processEntriesToToggl()
  } catch (error) {
    console.log('File not found or it\'s incorrect format')
  }
}

function main() {
  readFromExcel()
}

if (!apiKey || !filePath) {
  console.log('API key and file path to excel file required')
} else {
  main()
}
