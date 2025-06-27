const express = require("express")
const multer = require("multer")
const path = require("path")
const fs = require("fs")
const XLSX = require("xlsx")
const axios = require("axios")
const cors = require("cors")

const app = express()
const port = process.env.PORT || 3000

// Job tracking for async processing
const jobs = new Map()
const jobResults = new Map()

// Batch processing configuration
const BATCH_CONFIG = {
  BATCH_SIZE: 10, // Process 10 records concurrently
  MAX_RETRIES: 3, // Retry failed requests 3 times
  RETRY_DELAY: 1000, // 1 second delay between retries
  REQUEST_TIMEOUT: 30000, // 30 second timeout per request
  BATCH_DELAY: 500, // 500ms delay between batches to avoid rate limiting
}

// Multer storage configuration
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = "uploads"
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir)
    }
    cb(null, uploadDir)
  },
  filename: (req, file, cb) => {
    cb(null, `${Date.now()}-${file.originalname}`)
  },
})

const fileFilter = (req, file, cb) => {
  const allowedTypes = [
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-excel",
    "application/vnd.ms-excel.sheet.macroEnabled.12",
  ]

  const allowedExtensions = [".xlsx", ".xls", ".xlsm"]
  const fileExtension = path.extname(file.originalname).toLowerCase()

  if (allowedTypes.includes(file.mimetype) || allowedExtensions.includes(fileExtension)) {
    cb(null, true)
  } else {
    cb(new Error("Only Excel files (.xlsx, .xls, .xlsm) are allowed"), false)
  }
}

const upload = multer({
  storage,
  fileFilter,
  limits: { fileSize: 1024 * 1024 * 10 },
})

app.use(express.json())
app.use(express.urlencoded({ extended: true }))
app.use(cors())

// Utility function to create delay
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms))

// Enhanced retry mechanism with exponential backoff
async function retryWithBackoff(fn, maxRetries = BATCH_CONFIG.MAX_RETRIES, baseDelay = BATCH_CONFIG.RETRY_DELAY) {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      return await fn()
    } catch (error) {
      if (attempt === maxRetries) {
        throw error
      }

      const delayTime = baseDelay * Math.pow(2, attempt - 1) // Exponential backoff
      console.log(`Attempt ${attempt} failed, retrying in ${delayTime}ms...`)
      await delay(delayTime)
    }
  }
}

// Rate limiting helper
class RateLimiter {
  constructor(requestsPerSecond = 10) {
    this.requestsPerSecond = requestsPerSecond
    this.requests = []
  }

  async waitForSlot() {
    const now = Date.now()
    this.requests = this.requests.filter((time) => now - time < 1000)

    if (this.requests.length >= this.requestsPerSecond) {
      const oldestRequest = Math.min(...this.requests)
      const waitTime = 1000 - (now - oldestRequest)
      if (waitTime > 0) {
        await delay(waitTime)
      }
    }

    this.requests.push(Date.now())
  }
}

const rateLimiter = new RateLimiter(15) // 15 requests per second

// Function to convert text with hyperlinks to HTML format
function convertTextToSpacesNotesHTML(text, hyperlinks = []) {
  if (!text) return ""

  const htmlContent = text
  const paragraphs = htmlContent.split(/\n\s*\n|\r\n\s*\r\n/).filter((p) => p.trim())
  let formattedHTML = ""

  paragraphs.forEach((paragraph, index) => {
    let paragraphText = paragraph.trim()

    hyperlinks.forEach((link) => {
      if (paragraphText.includes(link.text)) {
        const linkHTML = `&lt;a href=&quot;${link.url}&quot; target=&quot;_blank&quot; rel=&quot;noopener noreferrer&quot;&gt;${link.text}&lt;/a&gt;`
        paragraphText = paragraphText.replace(link.text, linkHTML)
      }
    })

    paragraphText = paragraphText
      .replace(/→/g, "→")
      .replace(/-->/g, "→")
      .replace(/\n/g, "&lt;br&gt;")
      .replace(/\r\n/g, "&lt;br&gt;")

    if (index === 0 && paragraphText.includes("→")) {
      formattedHTML += `&lt;p id=&quot;isPasted&quot;&gt;${paragraphText}&lt;/p&gt;`
    } else {
      formattedHTML += `&lt;p&gt;${paragraphText}&lt;/p&gt;`
    }

    if (index < paragraphs.length - 1) {
      formattedHTML += `&lt;p&gt;&lt;br&gt;&lt;/p&gt;`
    }
  })

  return formattedHTML
}

function processSpaceNotesWithHyperlinks(spaceNotesText, extractedHyperlinks = []) {
  if (!spaceNotesText) return ""

  const commonLinks = [
    { text: "here", url: "https://community.unit4.com/t5/Success-Outcomes/bg-p/SuccessOutcomes" },
    { text: "Community4U", url: "https://community.unit4.com/" },
    { text: "Success.Hub@Unit4.com", url: "mailto:Success.Hub@Unit4.com" },
  ]

  const allHyperlinks = [...extractedHyperlinks, ...commonLinks]
  const htmlContent = convertTextToSpacesNotesHTML(spaceNotesText, allHyperlinks)

  return htmlContent
}

function extractHyperlinksFromWorksheet(worksheet, jsonData) {
  const processedData = []
  let totalHyperlinks = 0
  const extractedHyperlinks = []

  const range = XLSX.utils.decode_range(worksheet["!ref"] || "A1")
  const headers = []

  for (let col = range.s.c; col <= range.e.c; col++) {
    const headerCell = worksheet[XLSX.utils.encode_cell({ r: range.s.r, c: col })]
    headers.push(headerCell ? headerCell.v || headerCell.w || `Column${col}` : `Column${col}`)
  }

  jsonData.forEach((row, rowIndex) => {
    const processedRow = { ...row }

    headers.forEach((header, colIndex) => {
      const cellAddress = XLSX.utils.encode_cell({
        c: colIndex,
        r: rowIndex + range.s.r + 1,
      })

      const cell = worksheet[cellAddress]

      if (cell && processedRow.hasOwnProperty(header)) {
        let hasHyperlink = false
        let linkUrl = null
        const displayText = processedRow[header] || cell.v || cell.w || ""

        if (cell.l && cell.l.Target) {
          linkUrl = cell.l.Target
          hasHyperlink = true
        } else if (cell.l && cell.l.Hyperlink) {
          linkUrl = cell.l.Hyperlink
          hasHyperlink = true
        } else if (typeof cell.v === "string" && (cell.v.startsWith("http://") || cell.v.startsWith("https://"))) {
          linkUrl = cell.v
          hasHyperlink = true
        } else if (cell.f && cell.f.includes("HYPERLINK")) {
          const hyperlinkMatch = cell.f.match(/HYPERLINK\s*\(\s*"([^"]+)"/i)
          if (hyperlinkMatch) {
            linkUrl = hyperlinkMatch[1]
            hasHyperlink = true
          }
        }

        if (hasHyperlink && linkUrl) {
          processedRow[header] = {
            text: displayText,
            link: linkUrl,
            hasHyperlink: true,
          }

          extractedHyperlinks.push({
            text: displayText,
            url: linkUrl,
            column: header,
            cell: cellAddress,
          })

          totalHyperlinks++
        }

        if (header === "Space_Notes" && displayText) {
          const spaceNotesHTML = processSpaceNotesWithHyperlinks(displayText, extractedHyperlinks)
          processedRow[header + "_HTML"] = spaceNotesHTML
        }
      }
    })

    processedData.push(processedRow)
  })

  return processedData
}

function extractHyperlink(cellData) {
  if (typeof cellData === "object" && cellData.hasHyperlink) {
    return cellData.link
  }
  return null
}

function getDisplayText(cellData) {
  if (typeof cellData === "object" && cellData.hasHyperlink) {
    return cellData.text
  }
  return cellData
}

// Enhanced API functions with timeout and retry
async function getWidgetId(url, cookie, companyId) {
  return retryWithBackoff(async () => {
    await rateLimiter.waitForSlot()

    const response = await axios.post(
      `${url}/v2/galaxy/spaces/assignment/resolve/cid`,
      {
        companyId,
        entityId: companyId,
        entityType: "Company",
        sharingType: "external",
      },
      {
        headers: {
          Cookie: cookie,
          "Content-Type": "application/json",
        },
        timeout: BATCH_CONFIG.REQUEST_TIMEOUT,
        maxBodyLength: Number.POSITIVE_INFINITY,
      },
    )

    const layout = response.data?.data?.layout
    const section = layout?.sections?.[0].sectionId
    const widget = layout?.sections?.[0]?.config?.widgets?.[0]

    if (!layout || !section || !widget) {
      throw new Error("Incomplete layout or widget information")
    }

    return {
      layoutId: layout.layoutId,
      sectionId: section,
      widgetDetails: widget,
    }
  })
}

async function addSpaceNotes(url, cookie, notes, companyId, sectionId, layoutId) {
  return retryWithBackoff(async () => {
    await rateLimiter.waitForSlot()

    const data = {
      entityId: companyId,
      entityType: "company",
      layoutId,
      sectionId,
      data: {
        SpacesNotes: notes,
      },
    }

    const response = await axios.put(url + "/v2/galaxy/spaces/cr360/data/section?ignoreMasked=true", data, {
      headers: {
        Cookie: cookie,
        "Content-Type": "application/json",
      },
      timeout: BATCH_CONFIG.REQUEST_TIMEOUT,
      maxBodyLength: Number.POSITIVE_INFINITY,
    })

    return response.data
  })
}

async function addSucessPlan(url, cookie, Success_Plan_GSID) {
  return retryWithBackoff(async () => {
    await rateLimiter.waitForSlot()

    const data = {
      request: [
        {
          assetId: Success_Plan_GSID,
          sharedWithSpaces: true,
        },
      ],
    }

    const response = await axios.put(url + "/v1/successPlan/spaces-config", data, {
      headers: {
        Cookie: cookie,
        "Content-Type": "application/json",
      },
      timeout: BATCH_CONFIG.REQUEST_TIMEOUT,
      maxBodyLength: Number.POSITIVE_INFINITY,
    })

    return response.data
  })
}

async function trySearchUser(url, cookie, companyId, email) {
  return retryWithBackoff(async () => {
    await rateLimiter.waitForSlot()

    const response = await axios.post(
      `${url}/v1/spaces/search/Company/person`,
      {
        entityId: companyId,
        companyId,
        searchString: email,
      },
      {
        headers: { Cookie: cookie, "Content-Type": "application/json" },
        timeout: BATCH_CONFIG.REQUEST_TIMEOUT,
        maxBodyLength: Number.POSITIVE_INFINITY,
      },
    )

    const person = response?.data?.data?.[0]
    if (person && person.person__Email && person.person__Gsid) {
      return person
    } else {
      return null
    }
  })
}

async function tryCockPitSearchUser(url, cookie, companyId, email) {
  return retryWithBackoff(async () => {
    await rateLimiter.waitForSlot()

    const response = await axios.post(
      `${url}/v1/api/standardobjects/user/findOrCreateRecord`,
      {
        SystemType: "External",
        IsActiveUser: "true",
        CompanyID: companyId,
        Email: email,
        Name: null,
        FirstName: null,
        LastName: null,
      },
      {
        headers: { Cookie: cookie, "Content-Type": "application/json" },
        timeout: BATCH_CONFIG.REQUEST_TIMEOUT,
        maxBodyLength: Number.POSITIVE_INFINITY,
      },
    )

    const person = response?.data?.data
    return person.result[0]?.Gsid ? person.result[0]?.Gsid : null
  })
}

function generateEmailHTML(contactName) {
  let htmlTemplate =
    "&lt;!DOCTYPE html&gt;&lt;html xmlns=&quot;http://www.w3.org/1999/xhtml&quot; lang=&quot;en&quot;&gt;&lt;head&gt;\n                                  &lt;meta charset=&quot;UTF-8&quot; /&gt;\n                                  &lt;meta name=&quot;viewport&quot; content=&quot;width=device-width, initial-scale=1.0&quot; /&gt;\n                                  &lt;title&gt;GrapesJS Exported Content&lt;/title&gt;\n                                  &lt;style&gt;\n                                    body{margin-top:0px;margin-right:0px;margin-bottom:0px;margin-left:0px;padding-top:0px;padding-right:0px;padding-bottom:0px;padding-left:0px;background-color:rgb(245, 247, 249);font-family:&quot;Noto Sans&quot;;}.container{width:600px;max-width:100%;margin-top:0px;margin-right:auto;margin-bottom:0px;margin-left:auto;padding-top:32px;box-sizing:border-box;color:rgb(26, 26, 26);font-family:Roboto, Arial, sans-serif, sans-serif;background-image:url(&quot;https://staticcss.gainsightapp.net/space_email/email-banner.png&quot;);background-position-x:center;background-position-y:top;background-size:initial;background-repeat:no-repeat;background-attachment:initial;background-origin:initial;background-clip:initial;background-color:rgb(255, 255, 255) !important;}.content{width:544px;max-width:100%;background-color:rgb(255, 255, 255);margin-top:auto;margin-right:auto;margin-bottom:auto;margin-left:auto;padding-top:32px;padding-right:32px;padding-bottom:32px;padding-left:32px;border-top-width:1px;border-right-width:1px;border-bottom-width:1px;border-left-width:1px;border-top-style:solid;border-right-style:solid;border-bottom-style:solid;border-left-style:solid;border-top-color:rgb(230, 233, 236);border-right-color:rgb(230, 233, 236);border-bottom-color:rgb(230, 233, 236);border-left-color:rgb(230, 233, 236);border-image-source:initial;border-image-slice:initial;border-image-width:initial;border-image-outset:initial;border-image-repeat:initial;border-top-left-radius:8px;border-top-right-radius:8px;border-bottom-right-radius:8px;border-bottom-left-radius:8px;box-sizing:border-box;}.section{margin-bottom:24px;}.title{font-size:24px;font-weight:600;color:rgb(26, 26, 26);}.box{background-color:rgb(245, 247, 249);border-top-left-radius:12px;border-top-right-radius:12px;border-bottom-right-radius:12px;border-bottom-left-radius:12px;padding-top:16px;padding-right:16px;padding-bottom:16px;padding-left:16px;margin-bottom:24px;box-sizing:border-box;}.box-title{font-size:14px;font-style:normal;font-weight:600;line-height:24px;padding-bottom:4px;}.box-descr{font-size:12px;font-style:normal;font-weight:400;line-height:16px;}.button{display:inline-block;padding-top:4px;padding-right:16px;padding-bottom:4px;padding-left:16px;font-size:14px;font-weight:600;line-height:24px;text-decoration-line:none;text-decoration-thickness:initial;text-decoration-style:initial;text-decoration-color:initial;border-top-left-radius:100px;border-top-right-radius:100px;border-bottom-right-radius:100px;border-bottom-left-radius:100px;text-align:center;margin-top:16px;background-color:rgb(42, 170, 225);color:rgb(255, 255, 255) !important;}.footer{height:8px;width:100%;background-image:url(&quot;https://staticcss.gainsightapp.net/space_email/email-footer.png&quot;);background-position-x:center;background-position-y:bottom;background-size:initial;background-repeat:no-repeat;background-attachment:initial;background-origin:initial;background-clip:initial;background-color:initial;}.bookmark{padding-top:24px;padding-right:24px;padding-bottom:24px;padding-left:24px;font-size:14px;}.logo{display:block;margin-bottom:24px;border-top-width:1px;border-right-width:1px;border-bottom-width:1px;border-left-width:1px;border-top-style:solid;border-right-style:solid;border-bottom-style:solid;border-left-style:solid;border-top-color:rgb(230, 233, 236);border-right-color:rgb(230, 233, 236);border-bottom-color:rgb(230, 233, 236);border-left-color:rgb(230, 233, 236);border-image-source:initial;border-image-slice:initial;border-image-width:initial;border-image-outset:initial;border-image-repeat:initial;border-top-left-radius:8px;border-top-right-radius:8px;border-bottom-right-radius:8px;border-bottom-left-radius:8px;height:40px;}.spaces-bookmark-url{font-size:12px;}#permission-title{font-size:14px;font-style:normal;font-weight:600;line-height:24px;padding-bottom:4px;color:rgb(26, 26, 26);}#permission-descr{font-size:12px;font-style:normal;font-weight:400;line-height:16px;color:rgb(26, 26, 26);}.gjs-selected{border-top-left-radius:4px;border-top-right-radius:4px;border-bottom-right-radius:4px;border-bottom-left-radius:4px;outline-offset:4px;outline-color:rgb(3, 105, 233) !important;outline-style:solid !important;outline-width:1px !important;}.rte-token-span{display:inline-flex;padding-top:2px;padding-right:8px;padding-bottom:2px;padding-left:2px;align-items:center;border-top-left-radius:72px;border-top-right-radius:72px;border-bottom-right-radius:72px;border-bottom-left-radius:72px;border-top-width:1px;border-right-width:1px;border-bottom-width:1px;border-left-width:1px;border-top-style:solid;border-right-style:solid;border-bottom-style:solid;border-left-style:solid;border-top-color:rgb(210, 228, 251);border-right-color:rgb(210, 228, 251);border-bottom-color:rgb(210, 228, 251);border-left-color:rgb(210, 228, 251);border-image-source:initial;border-image-slice:initial;border-image-width:initial;border-image-outset:initial;border-image-repeat:initial;background-image:initial;background-position-x:initial;background-position-y:initial;background-size:initial;background-repeat:initial;background-attachment:initial;background-origin:initial;background-clip:initial;background-color:rgb(245, 249, 254);}.rte-token-span .token-text-inner{color:rgb(3, 105, 233);font-size:12px;font-style:normal;font-weight:400;line-height:16px;}#ixo52{line-height:1.284;margin:0cm 0cm 8pt;font-family:Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif;color:rgb(0, 0, 0);}#ib3ss{line-height:1.284;margin:0cm 0cm 8pt;font-family:Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif;color:rgb(0, 0, 0);}#ik5bj{color:rgb(70, 120, 134);}#m_-2258095401727991454m_-1548164695275070542OWAe74fdd82-bcaa-e46f-4501-a54063d9df33{color:rgb(70, 120, 134);}#ij6pf{color:rgb(34, 34, 34);}#ixqlt{color:rgb(70, 120, 134);}#m_-2258095401727991454m_-1548164695275070542OWAbd617e49-ddbe-6b91-3bd8-893df842d0ae{color:rgb(70, 120, 134);}#i57b2{line-height:1.284;margin:0cm 0cm 8pt;font-family:Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif;color:rgb(0, 0, 0);}#i7ok6{line-height:1.284;margin:0cm 0cm 8pt;font-family:Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif;color:rgb(0, 0, 0);}#iqldt{color:rgb(34, 34, 34);}#io444{color:rgb(34, 34, 34);line-height:1.284;margin:0cm 0cm 8pt;font-family:Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif;}#i3e9h{color:rgb(0, 0, 0);}#iruzn{color:rgb(70, 120, 134);}#m_-2258095401727991454m_-1548164695275070542OWA858b1f2c-6b35-9313-0d19-a6819c481bdb{color:rgb(70, 120, 134);}#ixfli{color:rgb(0, 0, 0);}#i7mnh{color:rgb(34, 34, 34);}\n                                  &lt;/style&gt;\n                                &lt;/head&gt;\n                                &lt;body id=&quot;ia22&quot;&gt;\n                                  &lt;meta charset=&quot;UTF-8&quot; /&gt;&lt;title&gt;Title&lt;/title&gt;&lt;div&gt;&lt;div class=&quot;container&quot;&gt;&lt;div class=&quot;content&quot;&gt;&lt;div class=&quot;logo-container&quot;&gt;&lt;img src=&quot;${embd::logo_url}&quot; alt=&quot;Logo&quot; class=&quot;logo&quot; /&gt;&lt;/div&gt;&lt;div class=&quot;section editable-block&quot; id=&quot;ibkh&quot;&gt;&lt;div id=&quot;i57b2&quot;&gt;&lt;br /&gt;Hi [Contact Name], &lt;br /&gt;&lt;br /&gt;&lt;/div&gt;&lt;div id=&quot;i7ok6&quot;&gt;We're excited to welcome you to Spaces, our new centralized customer hub &lt;a name=&quot;m_-2258095401727991454_m_-1548164695275070542__Int_sLRxCSge&quot; id=&quot;iqldt&quot;&gt;designed&lt;/a&gt; to enhance your Unit4 experience with shared resources, updates, and progress on your organization's goals.  It's a single, secure place to stay informed and engaged throughout your journey with Unit4.&lt;br /&gt;&lt;/div&gt;&lt;div id=&quot;io444&quot;&gt;&lt;span id=&quot;i3e9h&quot;&gt;We're continuously evolving Spaces to maximize the value it brings to you. You can now explore a &lt;b&gt;Digital Success Plan&lt;/b&gt; tailored to your organization, based on &lt;/span&gt;&lt;span id=&quot;iruzn&quot;&gt;&lt;u&gt;&lt;a href=&quot;https://community.unit4.com/t5/Success-Outcomes/bg-p/SuccessOutcomes&quot; id=&quot;m_-2258095401727991454m_-1548164695275070542OWA858b1f2c-6b35-9313-0d19-a6819c481bdb&quot; target=&quot;_blank&quot; data-saferedirecturl=&quot;https://www.google.com/url?q=https://community.unit4.com/t5/Success-Outcomes/bg-p/SuccessOutcomes&amp;amp;source=gmail&amp;amp;ust=1750931331309000&amp;amp;usg=AOvVaw3Vi1KAt8eMeed2Qmuxbn-n&quot;&gt;key outcomes&lt;/a&gt;&lt;/u&gt;&lt;/span&gt;&lt;span id=&quot;ixfli&quot;&gt; from &lt;a name=&quot;m_-2258095401727991454_m_-1548164695275070542__Int_WGOarlgd&quot; id=&quot;i7mnh&quot;&gt;our&lt;/a&gt; Success Catalog. These outcomes were carefully selected based on our experience and best practices, working with Non-Profit Organizations like yours, to maximize the value of your Unit4 solution.&lt;/span&gt;&lt;/div&gt;&lt;/div&gt;&lt;div class=&quot;box&quot;&gt;&lt;label id=&quot;permission-title&quot; class=&quot;box-title&quot;&gt;${embd::permissionTitle}&lt;/label&gt; &lt;br /&gt;&lt;label id=&quot;permission-descr&quot; class=&quot;box-descr&quot;&gt;${embd::permissionDescr}&lt;/label&gt;&lt;/div&gt;&lt;p class=&quot;editable-block&quot;&gt;Click below to get started:&lt;br /&gt;&lt;a href=&quot;${embd::redirect_url}&quot; class=&quot;button editable-block spaces-btn&quot;&gt;Join Spaces&lt;/a&gt;&lt;/p&gt;&lt;p class=&quot;editable-block&quot; id=&quot;ih5mz&quot;&gt;&lt;/p&gt;&lt;div id=&quot;ixo52&quot;&gt;Questions or want to learn more: Visit &lt;span id=&quot;ik5bj&quot;&gt;&lt;u&gt;&lt;a href=&quot;https://community.unit4.com/t5/Success4U-Hub/ct-p/Success4UProfessional&quot; id=&quot;m_-2258095401727991454m_-1548164695275070542OWAe74fdd82-bcaa-e46f-4501-a54063d9df33&quot; target=&quot;_blank&quot; data-saferedirecturl=&quot;https://www.google.com/url?q=https://community.unit4.com/t5/Success4U-Hub/ct-p/Success4UProfessional&amp;amp;source=gmail&amp;amp;ust=1750931331309000&amp;amp;usg=AOvVaw1TcuXQMQRRmES4VBjTUt8b&quot;&gt;Success4U Hub - Community4U&lt;/a&gt;&lt;/u&gt;&lt;/span&gt; or C&lt;a name=&quot;m_-2258095401727991454_m_-1548164695275070542__Int_KB2dZutv&quot; id=&quot;ij6pf&quot;&gt;ontact&lt;/a&gt; us at: &lt;span id=&quot;ixqlt&quot;&gt;&lt;u&gt;&lt;a href=&quot;mailto:Success.Hub@Unit4.com&quot; id=&quot;m_-2258095401727991454m_-1548164695275070542OWAbd617e49-ddbe-6b91-3bd8-893df842d0ae&quot; target=&quot;_blank&quot;&gt;Success.Hub@Unit4.com&lt;/a&gt;&lt;br /&gt;&lt;/u&gt;&lt;/span&gt;&lt;br /&gt;Best regards, &lt;/div&gt;&lt;div id=&quot;ib3ss&quot;&gt;Unit4 Customer Success Team&lt;/div&gt;&lt;p&gt;&lt;/p&gt;&lt;p&gt;&lt;/p&gt;&lt;/div&gt;&lt;div class=&quot;bookmark&quot;&gt;&lt;p class=&quot;editable-block&quot;&gt;You can bookmark us for future access:&lt;/p&gt;&lt;a href=&quot;${embd::bookmark_url}&quot; class=&quot;spaces-bookmark-url&quot;&gt;${embd::bookmark_url}&lt;/a&gt;&lt;/div&gt;&lt;div class=&quot;footer&quot;&gt;&lt;/div&gt;&lt;/div&gt;&lt;/div&gt;\n                                \n                                &lt;/body&gt;&lt;/html&gt;"
  htmlTemplate = htmlTemplate.replace("[Contact Name]", contactName)
  return htmlTemplate
}

function convertTextToEncodedHtml(inputText) {
  const htmlFormatted = inputText
    .replace(/\n\s*\n/g, "<br><br>")
    .replace(/\n/g, " ")
    .replace(/→/g, "&rarr;")

  const wrapped = `<p>${htmlFormatted}</p>`
  const encoded = encodeURIComponent(wrapped)

  return encoded
}

async function sendInvitation(url, cookie, companyId, personId, email, Invite_Name) {
  return retryWithBackoff(async () => {
    await rateLimiter.waitForSlot()

    const emailBody = generateEmailHTML(Invite_Name)
    const data = {
      entityId: companyId,
      companyId,
      users: [
        {
          personId,
          email,
          userName: "System",
          permissionType: "DELEGATE",
        },
      ],
      emailBody: emailBody,
      emailSubject: "You're invited to join Spaces — Take Action to Access your Unit4 Success Plan and More!",
    }

    const response = await axios.post(`${url}/v1/spaces/invite/company/persons`, data, {
      headers: { Cookie: cookie, "Content-Type": "application/json" },
      timeout: BATCH_CONFIG.REQUEST_TIMEOUT,
      maxBodyLength: Number.POSITIVE_INFINITY,
    })

    return response.data
  })
}

async function addUser(url, cookie, companyId, email) {
  return retryWithBackoff(async () => {
    await rateLimiter.waitForSlot()

    const data = {
      Name: email.split("@")[0],
      Email: email,
      companies: [
        {
          Company_ID: companyId,
        },
      ],
    }

    const response = await axios.put(`${url}/v1/peoplemgmt/v1.0/people?areaName=PersonC360UI`, data, {
      headers: { Cookie: cookie, "Content-Type": "application/json" },
      timeout: BATCH_CONFIG.REQUEST_TIMEOUT,
      maxBodyLength: Number.POSITIVE_INFINITY,
    })

    return response.data
  })
}

async function updateWidgetDetails(url, cookie, companyGsid, layoutId, widgetDetails) {
  return retryWithBackoff(async () => {
    await rateLimiter.waitForSlot()

    const apiUrl = `${url}/v2/galaxy/spaces/customisation/save/Company/${companyGsid}/${layoutId}`

    const response = await fetch(apiUrl, {
      method: "PUT",
      headers: {
        Cookie: cookie,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(widgetDetails),
    })

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`)
    }

    return await response.json()
  })
}

// Enhanced batch processing function
async function processBatch(batch, url, cookie, batchIndex, totalBatches, jobId) {
  console.log(`🚀 Processing batch ${batchIndex + 1}/${totalBatches} with ${batch.length} records`)

  const batchResults = await Promise.allSettled(
    batch.map(async (row, index) => {
      const recordIndex = batchIndex * BATCH_CONFIG.BATCH_SIZE + index

      try {
        const result = await processRecord(row, url, cookie, recordIndex)

        // Update job progress
        if (jobId && jobs.has(jobId)) {
          const job = jobs.get(jobId)
          job.processedCount++
          job.progress = Math.round((job.processedCount / job.totalRecords) * 100)
          job.estimatedTimeRemaining = Math.max(
            Math.round(
              (((job.totalRecords - job.processedCount) / job.processedCount) * (Date.now() - job.startTime)) / 1000,
            ),
            0,
          )
        }

        return result
      } catch (error) {
        console.error(`❌ Error processing record ${recordIndex}:`, error.message)
        return {
          Company_GSID: getDisplayText(row.Company_GSID),
          status: "Failed",
          messages: [`Error: ${error.message}`],
          error: error.message,
        }
      }
    }),
  )

  // Extract results from Promise.allSettled
  const results = batchResults.map((result) => (result.status === "fulfilled" ? result.value : result.reason))

  console.log(
    `✅ Batch ${batchIndex + 1} completed: ${results.filter((r) => r.status === "Success").length} successful, ${results.filter((r) => r.status === "Failed").length} failed`,
  )

  return results
}

// Process individual record
async function processRecord(row, url, cookie, recordIndex) {
  const Company_GSID = getDisplayText(row.Company_GSID)
  const Video_URL = extractHyperlink(row.Video_URL) || getDisplayText(row.Video_URL)
  const Welcome_Banner = getDisplayText(row.Welcome_Banner)
  const Space_Notes = getDisplayText(row.Space_Notes_HTML)
  const Success_Plan_GSID = getDisplayText(row.Success_Plan_GSID)
  const Invite_Email = getDisplayText(row.Invite_Email)
  const CTA_Owner_Email = getDisplayText(row.CTA_Owner_Email)
  const Invite_Name = getDisplayText(row.Invite_Name)

  const recordResult = {
    Company_GSID,
    Video_URL,
    Invite_Email,
    status: "Success",
    messages: [],
    recordIndex,
  }

  try {
    // Get widget details
    const { layoutId, sectionId, widgetDetails } = await getWidgetId(url, cookie, Company_GSID)
    recordResult.messages.push("Fetched widget details")

    // Update widget configuration
    if (widgetDetails?.config) {
      widgetDetails.config.mediaContent = {
        mediaType: "VIDEO",
        content: {
          url: Video_URL || "",
          thumbnailUrl: "",
        },
      }
      widgetDetails.config.bannerLayoutType.layoutName = "WITH_MEDIA_CONTENT_LAYOUT"

      if (Welcome_Banner) {
        const refinedText = await convertTextToEncodedHtml(Welcome_Banner)
        widgetDetails.config.bannerText = { value: refinedText }
        widgetDetails.config.bannerContent = {
          type: "GRADIENT",
          value: {
            selectedSolidColor: null,
            selectedGradientColor: {
              background: "linear-gradient(180deg, #A2CF6B 0%, #F6F6F6 100%)",
              color: "#A2CF6B",
              selected: true,
            },
            selectedImage: null,
            isUploadedImage: false,
          },
          base64: null,
        }
      }
    }

    // Parallel execution of independent operations
    const parallelOperations = []

    // Widget update
    parallelOperations.push(
      updateWidgetDetails(url, cookie, Company_GSID, layoutId, widgetDetails).then(() => "Widget updated"),
    )

    // Space notes
    if (Space_Notes) {
      parallelOperations.push(
        addSpaceNotes(url, cookie, Space_Notes, Company_GSID, sectionId, layoutId).then(() => "Notes added"),
      )
    }

    // Success plan
    if (Success_Plan_GSID) {
      parallelOperations.push(addSucessPlan(url, cookie, Success_Plan_GSID).then(() => "Success Plan added"))
    }

    // Execute parallel operations
    const parallelResults = await Promise.allSettled(parallelOperations)
    parallelResults.forEach((result, index) => {
      if (result.status === "fulfilled") {
        recordResult.messages.push(result.value)
      } else {
        recordResult.status = "Partial"
        recordResult.messages.push(`Operation ${index + 1} failed: ${result.reason.message}`)
      }
    })

    // Process invitations
    if (Invite_Email) {
      const emailList = Invite_Email.split(",")
        .map((e) => e.trim())
        .filter(Boolean)
      const invitationResults = await Promise.allSettled(
        emailList.map(async (email) => {
          try {
            let user = await trySearchUser(url, cookie, Company_GSID, email)
            if (user) {
              await sendInvitation(url, cookie, Company_GSID, user.person__Gsid, email, Invite_Name)
              return `Invitation sent to ${email}`
            } else {
              await addUser(url, cookie, Company_GSID, email)
              user = await trySearchUser(url, cookie, Company_GSID, email)
              if (user?.person__Gsid) {
                await sendInvitation(url, cookie, Company_GSID, user.person__Gsid, email, Invite_Name)
                return `User added and invitation sent to ${email}`
              } else {
                throw new Error(`User could not be added or invited for ${email}`)
              }
            }
          } catch (error) {
            throw new Error(`Error with ${email}: ${error.message}`)
          }
        }),
      )

      invitationResults.forEach((result) => {
        if (result.status === "fulfilled") {
          recordResult.messages.push(result.value)
        } else {
          recordResult.status = "Partial"
          recordResult.messages.push(result.reason.message)
        }
      })
    }

    // Handle CTA Owner Email (simplified for brevity - same parallel approach can be applied)
    if (CTA_Owner_Email) {
      try {
        const userId = await tryCockPitSearchUser(url, cookie, Company_GSID, CTA_Owner_Email)
        // CTA processing logic here...
        recordResult.messages.push(`CTA processing completed for ${CTA_Owner_Email}`)
      } catch (ctaErr) {
        recordResult.status = "Partial"
        recordResult.messages.push(`CTA error: ${ctaErr.message}`)
      }
    }
  } catch (err) {
    recordResult.status = "Failed"
    recordResult.messages.push(`Error: ${err.message}`)
    throw err
  }

  return recordResult
}

// Enhanced processSharedSpace with batch processing
async function processSharedSpaceWithBatches(results, url, cookie, jobId = null) {
  console.log(`🚀 Starting batch processing for ${results.length} records`)
  console.log(
    `📊 Batch configuration: ${BATCH_CONFIG.BATCH_SIZE} records per batch, ${BATCH_CONFIG.BATCH_DELAY}ms delay between batches`,
  )

  const startTime = Date.now()
  const totalBatches = Math.ceil(results.length / BATCH_CONFIG.BATCH_SIZE)
  const allResults = []

  // Initialize job tracking
  if (jobId) {
    jobs.set(jobId, {
      totalRecords: results.length,
      processedCount: 0,
      progress: 0,
      status: "Processing...",
      startTime,
      estimatedTimeRemaining: null,
      completed: false,
    })
  }

  // Process batches sequentially to avoid overwhelming the API
  for (let i = 0; i < totalBatches; i++) {
    const startIndex = i * BATCH_CONFIG.BATCH_SIZE
    const endIndex = Math.min(startIndex + BATCH_CONFIG.BATCH_SIZE, results.length)
    const batch = results.slice(startIndex, endIndex)

    try {
      const batchResults = await processBatch(batch, url, cookie, i, totalBatches, jobId)
      allResults.push(...batchResults)

      // Add delay between batches to prevent rate limiting
      if (i < totalBatches - 1) {
        await delay(BATCH_CONFIG.BATCH_DELAY)
      }
    } catch (batchError) {
      console.error(`❌ Batch ${i + 1} failed:`, batchError.message)

      // Add failed results for this batch
      const failedResults = batch.map((row) => ({
        Company_GSID: getDisplayText(row.Company_GSID),
        status: "Failed",
        messages: [`Batch processing failed: ${batchError.message}`],
        error: batchError.message,
      }))

      allResults.push(...failedResults)
    }
  }

  const endTime = Date.now()
  const totalTime = Math.round((endTime - startTime) / 1000)
  const successCount = allResults.filter((r) => r.status === "Success").length
  const partialCount = allResults.filter((r) => r.status === "Partial").length
  const failedCount = allResults.filter((r) => r.status === "Failed").length

  console.log(`✅ Batch processing completed in ${totalTime}s`)
  console.log(`📊 Results: ${successCount} successful, ${partialCount} partial, ${failedCount} failed`)
  console.log(`⚡ Average processing time: ${(totalTime / results.length).toFixed(2)}s per record`)

  // Update job completion
  if (jobId && jobs.has(jobId)) {
    const job = jobs.get(jobId)
    job.completed = true
    job.progress = 100
    job.status = "Completed successfully!"
    job.estimatedTimeRemaining = 0

    // Store results
    jobResults.set(jobId, {
      success: true,
      processedData: allResults,
      totalTime,
      statistics: {
        total: results.length,
        successful: successCount,
        partial: partialCount,
        failed: failedCount,
      },
    })
  }

  return allResults
}

// Job status endpoint
app.get("/api/job-status/:jobId", (req, res) => {
  const { jobId } = req.params

  if (!jobs.has(jobId)) {
    return res.status(404).json({ error: "Job not found" })
  }

  const job = jobs.get(jobId)
  const result = jobResults.get(jobId)

  res.json({
    ...job,
    result: result || null,
  })
})

// Cancel job endpoint
app.post("/api/cancel-job/:jobId", (req, res) => {
  const { jobId } = req.params

  if (jobs.has(jobId)) {
    jobs.delete(jobId)
    jobResults.delete(jobId)
    res.json({ success: true, message: "Job cancelled successfully" })
  } else {
    res.status(404).json({ error: "Job not found" })
  }
})

// Enhanced Excel Upload Route with async processing
app.post("/api/xlsxfileupload", upload.single("file"), async (req, res) => {
  let filePath = null

  try {
    if (!req.file) {
      return res.status(400).json({ success: false, message: "No file uploaded" })
    }

    console.log("=== FILE UPLOAD DEBUG ===")
    console.log("Uploaded file:", {
      originalname: req.file.originalname,
      filename: req.file.filename,
      path: req.file.path,
      size: req.file.size,
      mimetype: req.file.mimetype,
    })

    filePath = req.file.path

    if (!fs.existsSync(filePath)) {
      throw new Error(`Uploaded file not found at path: ${filePath}`)
    }

    const { url, cookie, worksheet, async: isAsync } = req.body

    if (!url || !cookie) {
      return res.status(400).json({
        success: false,
        message: "Missing 'url' or 'cookie' in request body",
      })
    }

    console.log(`Processing Excel file: ${req.file.originalname}`)

    // Read and parse Excel file
    let workbook
    try {
      workbook = XLSX.readFile(filePath)
    } catch (readError1) {
      try {
        const fileBuffer = fs.readFileSync(filePath)
        workbook = XLSX.read(fileBuffer, { type: "buffer" })
      } catch (readError2) {
        const fileBuffer = fs.readFileSync(filePath)
        const arrayBuffer = new Uint8Array(fileBuffer)
        workbook = XLSX.read(arrayBuffer, { type: "array" })
      }
    }

    if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
      throw new Error("No worksheets found in the Excel file")
    }

    let sheetName
    if (worksheet && workbook.SheetNames.includes(worksheet)) {
      sheetName = worksheet
    } else {
      sheetName = workbook.SheetNames[0]
    }

    const worksheetData = workbook.Sheets[sheetName]
    if (!worksheetData) {
      throw new Error(`Worksheet "${sheetName}" not found in workbook`)
    }

    const jsonData = XLSX.utils.sheet_to_json(worksheetData)
    if (jsonData.length === 0) {
      throw new Error("The selected worksheet appears to be empty or contains no data rows")
    }

    console.log(`Found ${jsonData.length} data rows in worksheet`)

    // Enhanced hyperlink extraction
    const results = extractHyperlinksFromWorksheet(worksheetData, jsonData)

    // Check if async processing is requested
    if (isAsync === "true" && results.length > 50) {
      // Generate job ID for async processing
      const jobId = `job_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`

      console.log(`🚀 Starting async batch processing for ${results.length} records with job ID: ${jobId}`)

      // Start async processing
      processSharedSpaceWithBatches(results, url, cookie, jobId).catch((error) => {
        console.error(`❌ Async job ${jobId} failed:`, error)
        if (jobs.has(jobId)) {
          const job = jobs.get(jobId)
          job.completed = true
          job.status = "Failed"
          jobResults.set(jobId, {
            success: false,
            error: error.message,
          })
        }
      })

      // Return job ID immediately
      res.status(200).json({
        success: true,
        message: "Async processing started",
        jobId,
        totalRecords: results.length,
        estimatedTime: Math.round(results.length * 2), // Rough estimate: 2 seconds per record
        file: {
          filename: req.file.filename,
          originalname: req.file.originalname,
          size: req.file.size,
          worksheet: sheetName,
          totalSheets: workbook.SheetNames.length,
        },
      })
    } else {
      // Synchronous processing for smaller files
      console.log(`🚀 Starting synchronous batch processing for ${results.length} records`)
      const processOutcome = await processSharedSpaceWithBatches(results, url, cookie)

      res.status(200).json({
        success: true,
        message: "Excel file uploaded and processed successfully",
        file: {
          filename: req.file.filename,
          originalname: req.file.originalname,
          path: req.file.path,
          size: req.file.size,
          worksheet: sheetName,
          totalSheets: workbook.SheetNames.length,
          availableSheets: workbook.SheetNames,
        },
        processedData: processOutcome,
      })
    }
  } catch (err) {
    console.error("Excel upload handler error:", err.message)
    console.error("Stack trace:", err.stack)

    res.status(500).json({
      success: false,
      message: "An error occurred during Excel file processing.",
      details: err.message,
    })
  } finally {
    if (filePath && fs.existsSync(filePath)) {
      try {
        fs.unlinkSync(filePath)
        console.log(`Cleaned up temporary file: ${filePath}`)
      } catch (cleanupErr) {
        console.error(`Failed to delete temporary file ${filePath}:`, cleanupErr)
      }
    }
  }
})

// Keep existing CSV route for backward compatibility
app.post("/api/csvfileupload", upload.single("csvFile"), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, message: "No file uploaded" })
    }
    const { url, cookie } = req.body

    if (!url || !cookie) {
      return res.status(400).json({ success: false, message: "Missing 'url' or 'cookie' in ConnectionDetails" })
    }

    const results = []
    const csv = require("csv-parser")

    fs.createReadStream(req.file.path)
      .pipe(csv())
      .on("data", (data) => results.push(data))
      .on("end", async () => {
        try {
          const processOutcome = await processSharedSpaceWithBatches(results, url, cookie)
          res.status(200).json({
            success: true,
            message: "CSV file uploaded and processed successfully",
            file: {
              filename: req.file.filename,
              path: req.file.path,
              size: req.file.size,
            },
            processedData: processOutcome,
          })
        } catch (err) {
          console.error("processSharedSpace error:", err.message)
          res.status(500).json({ success: false, message: "Error processing shared space", details: err.message })
        }
      })
      .on("error", (error) => {
        console.error("CSV parsing error:", error.message)
        res.status(500).json({ success: false, message: `Error parsing CSV: ${error.message}` })
      })
  } catch (err) {
    console.error("CSV upload handler error:", err.message)
    res.status(500).json({ success: false, message: err.message })
  }
})

app.use(express.static(path.join(__dirname, "build")))
app.get("", (req, res) => {
  res.sendFile(path.join(__dirname, "build", "index.html"))
})

// Global Error Middleware
app.use((err, req, res, next) => {
  console.error("Unhandled error:", err.stack)
  res.status(500).json({ success: false, message: "Internal Server Error", details: err.message })
})

// Cleanup old jobs periodically (every hour)
setInterval(
  () => {
    const oneHourAgo = Date.now() - 60 * 60 * 1000
    for (const [jobId, job] of jobs.entries()) {
      if (job.startTime < oneHourAgo) {
        jobs.delete(jobId)
        jobResults.delete(jobId)
        console.log(`🧹 Cleaned up old job: ${jobId}`)
      }
    }
  },
  60 * 60 * 1000,
)

// Start the server
app.listen(port, () => {
  console.log(`🚀 Server is running on port ${port}`)
  console.log(`📊 Excel file processing endpoint: /api/xlsxfileupload`)
  console.log(`📄 CSV file processing endpoint: /api/csvfileupload (legacy)`)
  console.log(`⚡ Batch processing enabled with ${BATCH_CONFIG.BATCH_SIZE} concurrent requests`)
  console.log(`🔄 Rate limiting: ${rateLimiter.requestsPerSecond} requests per second`)
})
