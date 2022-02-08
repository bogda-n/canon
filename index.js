const axios = require('axios')
const cheerio = require('cheerio')
const fs = require('fs-extra')
const path = require('path')
const {default: PQueue} = require('p-queue')
const excelToJson = require('simple-excel-to-json')
const xlsx = require('xlsx')


const baseUrl = 'https://store.canon'


const fileName = 'canon.xlsx'
const file = excelToJson.parseXls2Json(path.resolve(__dirname, '_input', fileName))

const mpn = '1170C003'

const products = []

const locales = {
  'en': '.co.uk/',
  'it': '.it/',
  'fr': '.fr/',
  'de': '.de/'
}

const createReport = async () => {
  console.log('start create file report')
  const workBook = xlsx.utils.book_new()
  const workSheet = xlsx.utils.json_to_sheet(products)

  const outputDir = path.resolve(__dirname, '_outputs')
  await fs.ensureDir(outputDir)

  xlsx.utils.book_append_sheet(workBook, workSheet, 'Canon Matrix')
  xlsx.writeFile(workBook, `${outputDir}/result_${fileName}`)
}

const processing = async (sku) => {

  return new Promise(async (resolve, reject) => {
    try {
      const product = {}
      for (const locale in locales) {
        let blocks = {}
        const language = locales[locale]
        const url = `${baseUrl}${language}${sku.mpn}/`
        // console.log('processing', sku.mpn)
        try {
          const response = await axios.get(url)
          const $ = cheerio.load(response.data)

          $('div.page').each(async (idx, element) => {

            const head = $(element).find('h1 .pt_bold').text()
              .replace(/Canon/gmi, '').trim()
              .replace(/#\S+/gmi, '').trim()

            const text = $(element).find('[class="header-5 mom-tab--heading"]').text()

            const image = 'https:' + $(element)
              .find('[class="product-image-zoom"]')
              .find('a').attr('href')
              .replace(/\?w=\d+&/, '?w=8000&')

            product['RequestedBrandName'] = 'Canon'
            product['RequestedProductCode'] = sku.mpn
            product['ProductName INT'] = sku.mpn
            product['Image INT 1'] = image

            const translations = {
              benefits: ['benefits', 'points forts', 'vantaggi', 'vorteile'],
              package: ['what\'s in the box', 'contenu de la boîte', 'contenuto della confezione', 'lieferumfang'],
              compatibility: ['compatibility', 'compatibilité', 'compatibilità', 'kompatibilität']
            }

            $('.mom-tab--section').each(function () {
              const name = $(this).children('h4').text().trim()

              if (name) {
                const value = $(this).children('ul').text().replace(/^\s+/gm, '')

                if (translations.benefits.includes(name.toLowerCase())) {
                  blocks.benefits = { name, value }
                }

                if (translations.package.includes(name.toLowerCase())) {
                  blocks.package = { name, value }
                }

                if (translations.compatibility.includes(name.toLowerCase())) {
                  blocks.compatibility = { name, value }
                }
              }
            })

            if (locale === 'en') {
              product['ProductName EN'] = head
              product['LongProductName EN'] = head
              product['MarketingText EN'] = text
              product['BulletPoints EN'] = blocks.benefits?.value
              product['URL EN'] = url
              product['In the box EN'] = blocks.package?.value
            } else if (locale === 'fr') {
              product['ProductName FR'] = head
              product['LongProductName FR'] = head
              product['MarketingText FR'] = text
              product['BulletPoints FR'] = blocks.benefits?.value
              product['URL FR'] = url
              product['In the box FR'] = blocks.package?.value
            } else if (locale === 'de') {
              product['ProductName DE'] = head
              product['LongProductName DE'] = head
              product['MarketingText DE'] = text
              product['BulletPoints DE'] = blocks.benefits?.value
              product['URL DE'] = url
              product['In the box DE'] = blocks.package?.value
            } else if (locale === 'it') {
              product['ProductName IT'] = head
              product['LongProductName IT'] = head
              product['MarketingText IT'] = text
              product['BulletPoints IT'] = blocks.benefits?.value
              product['URL IT'] = url
              product['In the box IT'] = blocks.package?.value
            }

            product['Compatibility'] = blocks.compatibility?.value
          })
        } catch (e) {
          if (e.response?.status !== 404) {
            throw new Error(e)
          }
          console.error(e.message, '-' , sku.mpn)

        }
      }

      products.push(product)
      console.log('end processing', sku.mpn)
      resolve()

    } catch (e) {
      console.error(e.message)
      resolve()
    }
  })
}

const start = async () => {
  try {
    const queue = new PQueue({ concurrency: 5 })
    queue.on('add', () => {
      console.log(`Task is added.  Size: ${queue.size}  Pending: ${queue.pending}`)
    })
    queue.on('next', () => {
      console.log(`Task is completed.  Size: ${queue.size}  Pending: ${queue.pending}`)
    })
    queue.on('idle', () => {
      console.log('queue is clean', new Date())
      createReport()
    })
    // queue.add(() => Promise.resolve())
    for (let sku of file[0]) {
      queue.add(() => processing(sku))
    //   await processing(sku)
    }
  } catch (error) {
    console.error(error)
  }
}

start()