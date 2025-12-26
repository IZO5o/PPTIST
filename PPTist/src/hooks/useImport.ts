import { ref } from 'vue'
import { storeToRefs } from 'pinia'
import { parse, type Shape, type Element, type ChartItem, type BaseElement } from 'pptxtojson'
import { nanoid } from 'nanoid'
import { useSlidesStore } from '@/store'
import { decrypt } from '@/utils/crypto'
import { type ShapePoolItem, SHAPE_LIST, SHAPE_PATH_FORMULAS } from '@/configs/shapes'
import useAddSlidesOrElements from '@/hooks/useAddSlidesOrElements'
import useSlideHandler from '@/hooks/useSlideHandler'
import useHistorySnapshot from './useHistorySnapshot'
import message from '@/utils/message'
import { getSvgPathRange } from '@/utils/svgPathParser'
import type {
  Slide,
  SlideTheme,
  SlideType,
  PPTElement,
  PPTTextElement,
  TableCellStyle,
  TableCell,
  ChartType,
  SlideBackground,
  PPTShapeElement,
  PPTLineElement,
  PPTImageElement,
  ShapeTextAlign,
  ChartOptions,
  Gradient,
} from '@/types/slides'

const shapeVAlignMap: { [key: string]: ShapeTextAlign } = {
  'mid': 'middle',
  'down': 'bottom',
  'up': 'top',
}

const convertFontSizePtToPx = (html: string, ratio: number) => {
  return html.replace(/font-size:\s*([\d.]+)pt/g, (match, p1) => {
    return `font-size: ${(parseFloat(p1) * ratio).toFixed(1)}px`
  })
}

export default () => {
  const slidesStore = useSlidesStore()
  const { theme, viewportSize, viewportRatio } = storeToRefs(useSlidesStore())

  const { addHistorySnapshot } = useHistorySnapshot()
  const { addSlidesFromData } = useAddSlidesOrElements()
  const { isEmptySlide } = useSlideHandler()

  const exporting = ref(false)

  // 解析PPTX文件为PPTist可用的模板数据（不写入Store，不产生历史记录）
  const parsePPTXTemplate = async (files: FileList | File[], options?: { fixedViewport?: boolean }) => {
    const defaultOptions = {
      fixedViewport: true,
    }
    const { fixedViewport } = { ...defaultOptions, ...options }

    const file = files[0]
    if (!file) throw new Error('no file')

    const shapeList: ShapePoolItem[] = []
    for (const item of SHAPE_LIST) {
      shapeList.push(...item.children)
    }

    let json = null
    try {
      const buffer = await file.arrayBuffer()
      json = await parse(buffer)
    }
    catch {
      throw new Error('parse failed')
    }

    const baseTheme = (theme.value || {}) as SlideTheme
    const templateTheme: SlideTheme = {
      ...baseTheme,
      themeColors: json.themeColors,
    }

    let ratio = 96 / 72
    const width = json.size.width

    // 模板解析默认按当前画布宽度等比缩放，避免修改画布尺寸
    if (fixedViewport) ratio = viewportSize.value / width

    const slides: Slide[] = []
    for (const item of json.slides) {
      const { type, value } = item.fill
      let background: SlideBackground
      if (type === 'image') {
        background = {
          type: 'image',
          image: {
            src: value.picBase64,
            size: 'cover',
          },
        }
      }
      else if (type === 'gradient') {
        background = {
          type: 'gradient',
          gradient: {
            type: value.path === 'line' ? 'linear' : 'radial',
            colors: value.colors.map(item => ({
              ...item,
              pos: parseInt(item.pos),
            })),
            rotate: value.rot + 90,
          },
        }
      }
      else if (type === 'pattern') {
        background = {
          type: 'solid',
          color: '#fff',
        }
      }
      else {
        background = {
          type: 'solid',
          color: value || '#fff',
        }
      }

      const slide: Slide = {
        id: nanoid(10),
        elements: [],
        background,
        remark: item.note || '',
      }

      const defaultFontName = baseTheme.fontName
      const defaultFontColor = baseTheme.fontColor

      const parseElements = (elements: Element[]) => {
        const sortedElements = elements.sort((a, b) => a.order - b.order)

        for (const el of sortedElements) {
          const originWidth = el.width || 1
          const originHeight = el.height || 1
          const originLeft = el.left
          const originTop = el.top

          el.width = el.width * ratio
          el.height = el.height * ratio
          el.left = el.left * ratio
          el.top = el.top * ratio

          if (el.type === 'text') {
            if (el.autoFit && el.autoFit.type === 'text') {
              const fontScale = ratio * (el.autoFit.fontScale || 100) / 100
              const shapeEl: PPTShapeElement = {
                type: 'shape',
                id: nanoid(10),
                width: el.width,
                height: el.height,
                left: el.left,
                top: el.top,
                rotate: el.rotate,
                viewBox: [200, 200],
                path: 'M 0 0 L 200 0 L 200 200 L 0 200 Z',
                fill: el.fill.type === 'color' ? el.fill.value : '',
                fixedRatio: false,
                outline: {
                  color: el.borderColor,
                  width: +(el.borderWidth * ratio).toFixed(2),
                  style: el.borderType,
                },
                text: {
                  content: convertFontSizePtToPx(el.content, fontScale),
                  defaultFontName,
                  defaultColor: defaultFontColor,
                  align: shapeVAlignMap[el.vAlign] || 'middle',
                  lineHeight: 1,
                },
              }
              slide.elements.push(shapeEl)
            }
            else {
              const textEl: PPTTextElement = {
                type: 'text',
                id: nanoid(10),
                width: el.width,
                height: el.height,
                left: el.left,
                top: el.top,
                rotate: el.rotate,
                defaultFontName,
                defaultColor: defaultFontColor,
                content: convertFontSizePtToPx(el.content, ratio),
                lineHeight: 1,
                outline: {
                  color: el.borderColor,
                  width: +(el.borderWidth * ratio).toFixed(2),
                  style: el.borderType,
                },
                fill: el.fill.type === 'color' ? el.fill.value : '',
                vertical: el.isVertical,
              }
              if (el.shadow) {
                textEl.shadow = {
                  h: el.shadow.h * ratio,
                  v: el.shadow.v * ratio,
                  blur: el.shadow.blur * ratio,
                  color: el.shadow.color,
                }
              }
              slide.elements.push(textEl)
            }
          }
          else if (el.type === 'image') {
            const element: PPTImageElement = {
              type: 'image',
              id: nanoid(10),
              src: el.src,
              width: el.width,
              height: el.height,
              left: el.left,
              top: el.top,
              fixedRatio: true,
              rotate: el.rotate,
              flipH: el.isFlipH,
              flipV: el.isFlipV,
            }
            if (el.borderWidth) {
              element.outline = {
                color: el.borderColor,
                width: +(el.borderWidth * ratio).toFixed(2),
                style: el.borderType,
              }
            }
            const clipShapeTypes = ['roundRect', 'ellipse', 'triangle', 'rhombus', 'pentagon', 'hexagon', 'heptagon', 'octagon', 'parallelogram', 'trapezoid']
            if (el.rect) {
              element.clip = {
                shape: (el.geom && clipShapeTypes.includes(el.geom)) ? el.geom : 'rect',
                range: [
                  [
                    el.rect.l || 0,
                    el.rect.t || 0,
                  ],
                  [
                    100 - (el.rect.r || 0),
                    100 - (el.rect.b || 0),
                  ],
                ]
              }
            }
            else if (el.geom && clipShapeTypes.includes(el.geom)) {
              element.clip = {
                shape: el.geom,
                range: [[0, 0], [100, 100]]
              }
            }
            slide.elements.push(element)
          }
          else if (el.type === 'math') {
            slide.elements.push({
              type: 'image',
              id: nanoid(10),
              src: el.picBase64,
              width: el.width,
              height: el.height,
              left: el.left,
              top: el.top,
              fixedRatio: true,
              rotate: 0,
            })
          }
          else if (el.type === 'audio') {
            slide.elements.push({
              type: 'audio',
              id: nanoid(10),
              src: el.blob,
              width: el.width,
              height: el.height,
              left: el.left,
              top: el.top,
              rotate: 0,
              fixedRatio: false,
              color: templateTheme.themeColors?.[0] || baseTheme.themeColors?.[0] || '#000000',
              loop: false,
              autoplay: false,
            })
          }
          else if (el.type === 'video') {
            slide.elements.push({
              type: 'video',
              id: nanoid(10),
              src: (el.blob || el.src)!,
              width: el.width,
              height: el.height,
              left: el.left,
              top: el.top,
              rotate: 0,
              autoplay: false,
            })
          }
          else if (el.type === 'shape') {
            if (el.shapType === 'line' || /Connector/.test(el.shapType)) {
              const lineElement = parseLineElement(el, ratio)
              slide.elements.push(lineElement)
            }
            else {
              const shape = shapeList.find(item => item.pptxShapeType === el.shapType)

              const gradient: Gradient | undefined = el.fill?.type === 'gradient' ? {
                type: el.fill.value.path === 'line' ? 'linear' : 'radial',
                colors: el.fill.value.colors.map(item => ({
                  ...item,
                  pos: parseInt(item.pos),
                })),
                rotate: el.fill.value.rot,
              } : undefined

              const pattern: string | undefined = el.fill?.type === 'image' ? el.fill.value.picBase64 : undefined

              const fill = el.fill?.type === 'color' ? el.fill.value : ''

              const element: PPTShapeElement = {
                type: 'shape',
                id: nanoid(10),
                width: el.width,
                height: el.height,
                left: el.left,
                top: el.top,
                viewBox: [200, 200],
                path: 'M 0 0 L 200 0 L 200 200 L 0 200 Z',
                fill,
                gradient,
                pattern,
                fixedRatio: false,
                rotate: el.rotate,
                outline: {
                  color: el.borderColor,
                  width: +(el.borderWidth * ratio).toFixed(2),
                  style: el.borderType,
                },
                text: {
                  content: convertFontSizePtToPx(el.content, ratio),
                  defaultFontName,
                  defaultColor: defaultFontColor,
                  align: shapeVAlignMap[el.vAlign] || 'middle',
                },
                flipH: el.isFlipH,
                flipV: el.isFlipV,
              }
              if (el.shadow) {
                element.shadow = {
                  h: el.shadow.h * ratio,
                  v: el.shadow.v * ratio,
                  blur: el.shadow.blur * ratio,
                  color: el.shadow.color,
                }
              }

              if (shape) {
                element.path = shape.path
                element.viewBox = shape.viewBox

                if (shape.pathFormula) {
                  element.pathFormula = shape.pathFormula
                  element.viewBox = [el.width, el.height]

                  const pathFormula = SHAPE_PATH_FORMULAS[shape.pathFormula]
                  if ('editable' in pathFormula && pathFormula.editable) {
                    element.path = pathFormula.formula(el.width, el.height, pathFormula.defaultValue)
                    element.keypoints = pathFormula.defaultValue
                  }
                  else element.path = pathFormula.formula(el.width, el.height)
                }
              }
              else if (el.path && el.path.indexOf('NaN') === -1) {
                const { maxX, maxY } = getSvgPathRange(el.path)
                element.path = el.path
                if ((maxX / maxY) > (originWidth / originHeight)) {
                  element.viewBox = [maxX, maxX * originHeight / originWidth]
                }
                else {
                  element.viewBox = [maxY * originWidth / originHeight, maxY]
                }
              }
              if (el.shapType === 'custom') {
                if (el.path!.indexOf('NaN') !== -1) {
                  if (element.width === 0) element.width = 0.1
                  if (element.height === 0) element.height = 0.1
                  element.path = el.path!.replace(/NaN/g, '0')
                }
                else {
                  element.special = true
                  element.path = el.path!
                }
                const { maxX, maxY } = getSvgPathRange(element.path)
                if ((maxX / maxY) > (originWidth / originHeight)) {
                  element.viewBox = [maxX, maxX * originHeight / originWidth]
                }
                else {
                  element.viewBox = [maxY * originWidth / originHeight, maxY]
                }
              }

              if (element.path) slide.elements.push(element)
            }
          }
          else if (el.type === 'table') {
            const row = el.data.length
            const col = el.data[0].length

            const style: TableCellStyle = {
              fontname: defaultFontName,
              color: defaultFontColor,
            }
            const data: TableCell[][] = []
            for (let i = 0; i < row; i++) {
              const rowCells: TableCell[] = []
              for (let j = 0; j < col; j++) {
                const cellData = el.data[i][j]

                let textDiv: HTMLDivElement | null = document.createElement('div')
                textDiv.innerHTML = cellData.text
                const p = textDiv.querySelector('p')
                const align = p?.style.textAlign || 'left'

                const span = textDiv.querySelector('span')
                const fontsize = span?.style.fontSize ? (parseInt(span?.style.fontSize) * ratio).toFixed(1) + 'px' : ''
                const fontname = span?.style.fontFamily || ''
                const color = span?.style.color || cellData.fontColor

                rowCells.push({
                  id: nanoid(10),
                  colspan: cellData.colSpan || 1,
                  rowspan: cellData.rowSpan || 1,
                  text: textDiv.innerText,
                  style: {
                    ...style,
                    align: ['left', 'right', 'center'].includes(align) ? (align as 'left' | 'right' | 'center') : 'left',
                    fontsize,
                    fontname,
                    color,
                    bold: cellData.fontBold,
                    backcolor: cellData.fillColor,
                  },
                })
                textDiv = null
              }
              data.push(rowCells)
            }

            const allWidth = el.colWidths.reduce((a, b) => a + b, 0)
            const colWidths: number[] = el.colWidths.map(item => item / allWidth)

            const firstCell = el.data[0][0]
            const border = firstCell.borders.top ||
              firstCell.borders.bottom ||
              el.borders.top ||
              el.borders.bottom ||
              firstCell.borders.left ||
              firstCell.borders.right ||
              el.borders.left ||
              el.borders.right
            const borderWidth = border?.borderWidth || 0
            const borderStyle = border?.borderType || 'solid'
            const borderColor = border?.borderColor || '#eeece1'

            slide.elements.push({
              type: 'table',
              id: nanoid(10),
              width: el.width,
              height: el.height,
              left: el.left,
              top: el.top,
              colWidths,
              rotate: 0,
              data,
              outline: {
                width: +(borderWidth * ratio || 2).toFixed(2),
                style: borderStyle,
                color: borderColor,
              },
              cellMinHeight: el.rowHeights[0] ? el.rowHeights[0] * ratio : 36,
            })
          }
          else if (el.type === 'chart') {
            let labels: string[]
            let legends: string[]
            let series: number[][]

            if (el.chartType === 'scatterChart' || el.chartType === 'bubbleChart') {
              labels = el.data[0].map((item, index) => `坐标${index + 1}`)
              legends = ['X', 'Y']
              series = el.data
            }
            else {
              const data = el.data as ChartItem[]
              labels = Object.values(data[0].xlabels)
              legends = data.map(item => item.key)
              series = data.map(item => item.values.map(v => v.y))
            }

            const options: ChartOptions = {}

            let chartType: ChartType = 'bar'

            switch (el.chartType) {
              case 'barChart':
              case 'bar3DChart':
                chartType = 'bar'
                if (el.barDir === 'bar') chartType = 'column'
                if (el.grouping === 'stacked' || el.grouping === 'percentStacked') options.stack = true
                break
              case 'lineChart':
              case 'line3DChart':
                if (el.grouping === 'stacked' || el.grouping === 'percentStacked') options.stack = true
                chartType = 'line'
                break
              case 'areaChart':
              case 'area3DChart':
                if (el.grouping === 'stacked' || el.grouping === 'percentStacked') options.stack = true
                chartType = 'area'
                break
              case 'scatterChart':
              case 'bubbleChart':
                chartType = 'scatter'
                break
              case 'pieChart':
              case 'pie3DChart':
                chartType = 'pie'
                break
              case 'radarChart':
                chartType = 'radar'
                break
              case 'doughnutChart':
                chartType = 'ring'
                break
              default:
            }

            slide.elements.push({
              type: 'chart',
              id: nanoid(10),
              chartType: chartType,
              width: el.width,
              height: el.height,
              left: el.left,
              top: el.top,
              rotate: 0,
              themeColors: el.colors.length ? el.colors : (templateTheme.themeColors || []),
              textColor: defaultFontColor,
              data: {
                labels,
                legends,
                series,
              },
              options,
            })
          }
          else if (el.type === 'group') {
            let elements: BaseElement[] = el.elements.map(_el => {
              let left = _el.left + originLeft
              let top = _el.top + originTop

              if (el.rotate) {
                const { x, y } = calculateRotatedPosition(originLeft, originTop, originWidth, originHeight, _el.left, _el.top, el.rotate)
                left = x
                top = y
              }

              const element = {
                ..._el,
                left,
                top,
              }
              if (el.isFlipH && 'isFlipH' in element) element.isFlipH = true
              if (el.isFlipV && 'isFlipV' in element) element.isFlipV = true

              return element
            })
            if (el.isFlipH) elements = flipGroupElements(elements, 'y')
            if (el.isFlipV) elements = flipGroupElements(elements, 'x')
            parseElements(elements)
          }
          else if (el.type === 'diagram') {
            const elements = el.elements.map(_el => ({
              ..._el,
              left: _el.left + originLeft,
              top: _el.top + originTop,
            }))
            parseElements(elements)
          }
        }
      }

      parseElements([...item.elements, ...item.layoutElements])
      slides.push(slide)
    }

    const getHtmlText = (el: PPTElement): string => {
      if (el.type === 'text') return el.content || ''
      if (el.type === 'shape' && el.text) return el.text.content || ''
      return ''
    }

    const htmlToPlainText = (html: string) => {
      return html
        .replace(/<br\s*\/?\s*>/gi, '\n')
        .replace(/<[^>]*>/g, '')
        .replace(/\s+/g, ' ')
        .trim()
    }

    const getMaxFontSizePx = (html: string) => {
      const regex = /font-size:\s*(\d+(?:\.\d+)?)\s*px/gi
      let max = 0
      let match: RegExpExecArray | null
      while ((match = regex.exec(html))) {
        const val = parseFloat(match[1])
        if (!Number.isNaN(val)) max = Math.max(max, val)
      }
      return max || 16
    }

    const setTextType = (el: PPTElement, textType: any) => {
      if (el.type === 'text') {
        ;(el as any).textType = textType
        return
      }
      if (el.type === 'shape' && el.text) {
        ;(el.text as any).type = textType
      }
    }

    const clearTextType = (el: PPTElement) => {
      if (el.type === 'text') {
        delete (el as any).textType
        return
      }
      if (el.type === 'shape' && el.text) {
        delete (el.text as any).type
      }
    }

    const getTextCandidates = (slide: Slide): Array<PPTTextElement | PPTShapeElement> => {
      return slide.elements.filter(el => el.type === 'text' || (el.type === 'shape' && !!(el as PPTShapeElement).text)) as Array<PPTTextElement | PPTShapeElement>
    }

    const findByKeywords = (items: Array<{ el: PPTElement; plain: string }>, keywords: string[]) => {
      if (!keywords.length) return null
      const lowerKeywords = keywords.map(k => k.toLowerCase())
      const matched = items.filter(it => {
        const t = (it.plain || '').toLowerCase()
        return lowerKeywords.some(k => k && t.includes(k))
      })
      if (!matched.length) return null
      // 命中后仍按字号/位置择优
      return matched
        .map(it => {
          const font = getMaxFontSizePx(getHtmlText(it.el))
          return { ...it, font }
        })
        .sort((a, b) => (b.font - a.font) || ((a.el.top || 0) - (b.el.top || 0)))[0].el
    }

    const hasTextType = (slide: Slide, textType: string) => {
      return slide.elements.some(el => {
        if (el.type === 'text') return (el as any).textType === textType
        if (el.type === 'shape' && el.text) return (el.text as any).type === textType
        return false
      })
    }

    const slideHeight = viewportSize.value * viewportRatio.value

    const normalizePlaceholderFontColor = (color: string) => {
      const c = (color || '').trim().toLowerCase()
      if (!c) return '#333333'
      if (c === '#fff' || c === '#ffffff' || c === 'white') return '#333333'
      const rgb = c.match(/^rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/)
      if (rgb) {
        const r = +rgb[1]
        const g = +rgb[2]
        const b = +rgb[3]
        if (r >= 245 && g >= 245 && b >= 245) return '#333333'
      }
      return color
    }

    const isFooterLikeEl = (el: PPTElement) => {
      const top = (el as any).top || 0
      const h = (el as any).height || 0
      return top >= slideHeight * 0.9 && h <= Math.max(18, slideHeight * 0.04)
    }

    const isBrandLikeEl = (plain: string, el: PPTElement) => {
      const t = (plain || '').toLowerCase()
      if (!t) return false
      if (/(https?:\/\/|www\.|\.(com|cn|net|org)\b)/i.test(t)) return true
      if (/officeplus/i.test(t)) return true

      const left = (el as any).left || 0
      const top = (el as any).top || 0
      const w = (el as any).width || 0
      const h = (el as any).height || 0
      const font = getMaxFontSizePx(getHtmlText(el))
      const inTopRight = left >= viewportSize.value * 0.65 && top <= slideHeight * 0.18
      if (inTopRight && font <= 22 && h <= slideHeight * 0.08 && w <= viewportSize.value * 0.35) return true
      return false
    }

    const isTooTinyEl = (el: PPTElement, font: number) => {
      const w = (el as any).width || 0
      const h = (el as any).height || 0
      const area = w * h
      const slideArea = viewportSize.value * slideHeight
      if (area <= slideArea * 0.004 && font <= 12) return true
      if (w <= viewportSize.value * 0.12 && font <= 12) return true
      if (h <= 10) return true
      return false
    }

    const makeEmptyContentHtml = (fontSize: number, align: 'left' | 'center' | 'right' = 'left') => {
      // 使用 &nbsp; 保证存在 TextNode，便于后续 getNewTextElement 替换
      const color = normalizePlaceholderFontColor(baseTheme.fontColor)
      return `<p style="font-size: ${fontSize}px; font-family: ${baseTheme.fontName}; color: ${color}; text-align: ${align};">&nbsp;</p>`
    }

    const addTextBox = (slide: Slide, config: {
      left: number
      top: number
      width: number
      height: number
      fontSize: number
      align?: 'left' | 'center' | 'right'
      textType: string
      groupId?: string
    }) => {
      const el: PPTTextElement = {
        type: 'text',
        id: nanoid(10),
        left: config.left,
        top: config.top,
        width: config.width,
        height: config.height,
        rotate: 0,
        defaultFontName: baseTheme.fontName,
        defaultColor: baseTheme.fontColor,
        content: makeEmptyContentHtml(config.fontSize, config.align || 'left'),
        lineHeight: 1.2,
        textType: config.textType as any,
        groupId: config.groupId,
      }
      slide.elements.push(el)
    }

    const ensurePlaceholders = (slide: Slide, slideType: SlideType) => {
      // 仅在缺少关键占位符时补齐，避免覆盖用户 PPTX 原有布局
      // 对目录/内容页：如果 PPTX 自带文本框，则不要叠加固定字号的占位框，避免“字号不对/只剩目录标题”的观感
      // 注意：不能仅凭“存在文本框”就直接 return；否则一旦自动打标失败/误判，就永远不会补齐占位符。
      // 这里仅依据 textType 是否已成功打标来决定是否需要补齐。

      const marginX = Math.max(40, viewportSize.value * 0.06)
      const marginTop = Math.max(30, slideHeight * 0.06)
      const innerWidth = viewportSize.value - marginX * 2

      if (slideType === 'cover') {
        if (!hasTextType(slide, 'title')) {
          addTextBox(slide, {
            left: marginX,
            top: marginTop,
            width: innerWidth,
            height: slideHeight * 0.18,
            fontSize: 40,
            align: 'center',
            textType: 'title',
          })
        }
        if (!hasTextType(slide, 'content')) {
          addTextBox(slide, {
            left: marginX,
            top: marginTop + slideHeight * 0.2,
            width: innerWidth,
            height: slideHeight * 0.22,
            fontSize: 20,
            align: 'center',
            textType: 'content',
          })
        }
        return
      }

      if (slideType === 'transition') {
        if (!hasTextType(slide, 'title')) {
          addTextBox(slide, {
            left: marginX,
            top: slideHeight * 0.28,
            width: innerWidth,
            height: slideHeight * 0.16,
            fontSize: 36,
            align: 'center',
            textType: 'title',
          })
        }
        if (!hasTextType(slide, 'content')) {
          addTextBox(slide, {
            left: marginX,
            top: slideHeight * 0.46,
            width: innerWidth,
            height: slideHeight * 0.18,
            fontSize: 18,
            align: 'center',
            textType: 'content',
          })
        }
        if (!hasTextType(slide, 'partNumber')) {
          addTextBox(slide, {
            left: viewportSize.value - marginX - 120,
            top: marginTop,
            width: 120,
            height: 60,
            fontSize: 24,
            align: 'right',
            textType: 'partNumber',
          })
        }
        return
      }

      if (slideType === 'end') {
        // 结束页：尽量只准备一个居中的 title，占位用于后续替换为“谢谢/Thanks”
        if (!hasTextType(slide, 'title')) {
          addTextBox(slide, {
            left: marginX,
            top: slideHeight * 0.38,
            width: innerWidth,
            height: slideHeight * 0.18,
            fontSize: 40,
            align: 'center',
            textType: 'title',
          })
        }
        return
      }

      if (slideType === 'contents') {
        // 目录页：至少准备 10 行 item + itemNumber，便于适配拆分页
        const need = 10
        const startY = slideHeight * 0.22
        const rowH = Math.max(28, slideHeight * 0.055)
        const numberW = 60
        const textW = innerWidth - numberW

        // 如果 PPTX 自带 item/itemNumber 就不重复加
        const hasAnyItem = hasTextType(slide, 'item')
        const hasAnyNumber = hasTextType(slide, 'itemNumber')

        if (!hasAnyItem || !hasAnyNumber) {
          for (let i = 0; i < need; i++) {
            const gid = nanoid(10)
            if (!hasAnyNumber) {
              addTextBox(slide, {
                left: marginX,
                top: startY + i * rowH,
                width: numberW,
                height: rowH,
                fontSize: 18,
                align: 'right',
                textType: 'itemNumber',
                groupId: gid,
              })
            }
            if (!hasAnyItem) {
              addTextBox(slide, {
                left: marginX + numberW,
                top: startY + i * rowH,
                width: textW,
                height: rowH,
                fontSize: 18,
                align: 'left',
                textType: 'item',
                groupId: gid,
              })
            }
          }
        }
        // 目录也需要一个标题占位
        if (!hasTextType(slide, 'title')) {
          addTextBox(slide, {
            left: marginX,
            top: marginTop,
            width: innerWidth,
            height: slideHeight * 0.14,
            fontSize: 32,
            align: 'center',
            textType: 'title',
          })
        }
        return
      }

      if (slideType === 'content') {
        // 内容页必须有 title；并准备：单段 content + 4 组 itemTitle/item/itemNumber
        if (!hasTextType(slide, 'title')) {
          addTextBox(slide, {
            left: marginX,
            top: marginTop,
            width: innerWidth,
            height: slideHeight * 0.12,
            fontSize: 28,
            align: 'left',
            textType: 'title',
          })
        }

        if (!hasTextType(slide, 'content')) {
          addTextBox(slide, {
            left: marginX,
            top: slideHeight * 0.22,
            width: innerWidth,
            height: slideHeight * 0.60,
            fontSize: 18,
            align: 'left',
            textType: 'content',
          })
        }

        const need = 4
        const startY = slideHeight * 0.22
        const blockH = slideHeight * 0.15
        const numberW = 46
        const titleH = Math.max(26, blockH * 0.38)
        const textH = Math.max(34, blockH * 0.62)

        const hasAnyItem = hasTextType(slide, 'item')
        const hasAnyItemTitle = hasTextType(slide, 'itemTitle')
        const hasAnyNumber = hasTextType(slide, 'itemNumber')

        if (!hasAnyItem || !hasAnyItemTitle || !hasAnyNumber) {
          for (let i = 0; i < need; i++) {
            const gid = nanoid(10)
            const baseY = startY + i * blockH
            if (!hasAnyNumber) {
              addTextBox(slide, {
                left: marginX,
                top: baseY,
                width: numberW,
                height: titleH,
                fontSize: 18,
                align: 'right',
                textType: 'itemNumber',
                groupId: gid,
              })
            }
            if (!hasAnyItemTitle) {
              addTextBox(slide, {
                left: marginX + numberW + 10,
                top: baseY,
                width: innerWidth - numberW - 10,
                height: titleH,
                fontSize: 18,
                align: 'left',
                textType: 'itemTitle',
                groupId: gid,
              })
            }
            if (!hasAnyItem) {
              addTextBox(slide, {
                left: marginX + numberW + 10,
                top: baseY + titleH,
                width: innerWidth - numberW - 10,
                height: textH,
                fontSize: 16,
                align: 'left',
                textType: 'item',
                groupId: gid,
              })
            }
          }
        }
      }
    }

    const autoTagSlideTextTypes = (slide: Slide, slideType: SlideType) => {
      const candidates = getTextCandidates(slide)
      if (!candidates.length) return

      // 先清掉旧标记，避免误标导致后续缺字
      for (const el of candidates) clearTextType(el)

      const scored = candidates.map(el => {
        const html = getHtmlText(el)
        const font = getMaxFontSizePx(html)
        const area = (el.width || 0) * (el.height || 0)
        // 字体优先，其次靠上，其次面积
        const score = font * 100000 - (el.top || 0) * 10 + area / 100
        return { el, font, area, score, plain: htmlToPlainText(html) }
      }).sort((a, b) => b.score - a.score)

      const keywordTitleMap: Record<SlideType, string[]> = {
        cover: ['标题', '主题', 'title'],
        transition: ['章节', '部分', 'part', 'chapter', 'title'],
        contents: ['目录', 'contents', 'outline', 'agenda'],
        content: ['标题', 'title'],
        end: ['致谢', 'thanks', 'thank', '结束'],
      }

      const nonBrand = scored.filter(s => !isBrandLikeEl(s.plain, s.el))
      const titlePool = nonBrand.length ? nonBrand : scored
      const topHalf = titlePool.filter(s => (s.el.top || 0) <= slideHeight * 0.45)
      const fallbackTitleEl = (topHalf.length ? topHalf : titlePool)[0]?.el
      const kwTitleEl = findByKeywords(titlePool.map(s => ({ el: s.el, plain: s.plain })), keywordTitleMap[slideType] || [])
      const titleEl = kwTitleEl || fallbackTitleEl
      if (titleEl) setTextType(titleEl, 'title')

      const rest = scored.filter(s => s.el !== titleEl).filter(s => !isBrandLikeEl(s.plain, s.el))

      const slideArea = viewportSize.value * slideHeight

      const isFooterLike = (s: typeof rest[number]) => {
        return isFooterLikeEl(s.el)
      }

      const isTooTiny = (s: typeof rest[number]) => {
        // 很小的文本一般是装饰/页脚；保守起见不参与 AIPPT 填充
        return isTooTinyEl(s.el, s.font)
      }

      const isNumberBox = (s: typeof rest[number]) => {
        const plain = s.plain || ''
        if (/^\d+\.?$/.test(plain)) return true
        if ((s.el.width || 0) > 0 && (s.el.width || 0) < viewportSize.value * 0.12) return true
        return false
      }

      if (slideType === 'cover' || slideType === 'transition') {
        const kwContentEl = findByKeywords(rest.map(s => ({ el: s.el, plain: s.plain })), ['简介', '副标题', '说明', 'content', 'subtitle'])
        const fallbackContentEl = rest
          .slice()
          .sort((a, b) => (b.area - a.area) || ((a.el.top || 0) - (b.el.top || 0)))[0]?.el
        const contentEl = kwContentEl || fallbackContentEl
        if (contentEl) setTextType(contentEl, 'content')

        if (slideType === 'transition') {
          const numberEl = rest.filter(isNumberBox).sort((a, b) => ((a.el.top || 0) - (b.el.top || 0)) || ((a.el.left || 0) - (b.el.left || 0)))[0]?.el
          if (numberEl) setTextType(numberEl, 'partNumber')
        }
        return
      }

      if (slideType === 'contents') {
        const listRegion = rest
          .filter(s => (s.el.top || 0) >= slideHeight * 0.18 && (s.el.top || 0) <= slideHeight * 0.92)
          .filter(s => !isFooterLike(s) && !isTooTiny(s))

        const region = (listRegion.length ? listRegion : rest.filter(s => !isFooterLike(s) && !isTooTiny(s)))
        if (!region.length) return

        const ordered = region
          .slice()
          .sort((a, b) => ((a.el.top || 0) - (b.el.top || 0)) || ((a.el.left || 0) - (b.el.left || 0)))

        // 目录页保守打标：只选择“像列表项”的宽文本框；避免把装饰文字/页脚当成目录项
        const wideTextBoxes = ordered.filter(s => {
          const w = s.el.width || 0
          const h = s.el.height || 0
          const area = w * h
          if (w < viewportSize.value * 0.28) return false
          if (h > slideHeight * 0.14) return false
          if (area > slideArea * 0.25) return false
          return true
        })

        // 置信度不足：不打标，交给 ensurePlaceholders() 补齐标准占位符
        if (wideTextBoxes.length < 3) return

        const lefts = wideTextBoxes.map(s => s.el.left || 0).sort((a, b) => a - b)
        const itemLeft = lefts[Math.floor(lefts.length / 2)]
        const leftTol = Math.max(24, viewportSize.value * 0.06)

        const itemBoxes = wideTextBoxes
          .filter(s => Math.abs((s.el.left || 0) - itemLeft) <= leftTol)
          .sort((a, b) => ((a.el.top || 0) - (b.el.top || 0)))

        if (itemBoxes.length < 3) return

        for (const s of itemBoxes) setTextType(s.el, 'item')

        const narrowCandidates = ordered.filter(s => {
          const w = s.el.width || 0
          if (w > viewportSize.value * 0.18) return false
          if ((s.el.left || 0) >= itemLeft) return false
          return !isTooTiny(s)
        })

        // 尝试为每行匹配一个编号框（若匹配不到也不强行打标）
        for (const it of itemBoxes) {
          const itTop = it.el.top || 0
          const itH = it.el.height || 0
          const tolY = Math.max(16, Math.min(36, itH * 0.8))
          const matched = narrowCandidates
            .filter(n => Math.abs((n.el.top || 0) - itTop) <= tolY)
            .sort((a, b) => ((a.el.left || 0) - (b.el.left || 0)))[0]
          if (matched) setTextType(matched.el, 'itemNumber')
        }

        // 保底：至少要有 item
        return
      }

      if (slideType === 'content') {
        const bodyRegion = rest
          .filter(s => (s.el.top || 0) >= slideHeight * 0.16)
          .filter(s => !isFooterLike(s) && !isTooTiny(s))

        const region = (bodyRegion.length ? bodyRegion : rest.filter(s => !isFooterLike(s) && !isTooTiny(s)))
        if (!region.length) return

        const ordered = region
          .slice()
          .sort((a, b) => ((a.el.top || 0) - (b.el.top || 0)) || ((a.el.left || 0) - (b.el.left || 0)))

        // 编号框：只对较窄且靠左的框打标，避免把装饰小字误当编号
        const numberBoxes = ordered.filter(s => {
          const w = s.el.width || 0
          if (w > viewportSize.value * 0.14) return false
          if ((s.el.left || 0) > viewportSize.value * 0.5) return false
          return isNumberBox(s)
        })
        for (const s of numberBoxes) setTextType(s.el, 'itemNumber')

        const textBoxes = ordered.filter(s => !numberBoxes.includes(s))

        // 正文 content：优先选“面积明显更大且足够宽”的文本框；避免误选小装饰文本
        const contentCandidate = textBoxes
          .filter(s => {
            const w = s.el.width || 0
            const h = s.el.height || 0
            if (w < viewportSize.value * 0.35) return false
            if (h < slideHeight * 0.12) return false
            if (s.area < slideArea * 0.06) return false
            return true
          })
          .slice()
          .sort((a, b) => (b.area - a.area) || (b.font - a.font))[0]

        if (!contentCandidate) {
          // 置信度不足：不继续乱标注，交给 ensurePlaceholders() 补齐
          return
        }

        setTextType(contentCandidate.el, 'content')

        const remaining = textBoxes.filter(s => s.el !== contentCandidate.el)

        // 成对打标 itemTitle/item：只在存在“多组中等大小文本框”时才做
        const pairCandidates = remaining.filter(s => {
          const w = s.el.width || 0
          const h = s.el.height || 0
          if (w < viewportSize.value * 0.22) return false
          if (h < 14 || h > slideHeight * 0.18) return false
          if (s.area < slideArea * 0.01) return false
          return true
        }).sort((a, b) => ((a.el.top || 0) - (b.el.top || 0)) || ((a.el.left || 0) - (b.el.left || 0)))

        const maxPairs = Math.min(4, Math.floor(pairCandidates.length / 2))
        for (let i = 0; i < maxPairs; i++) {
          const a = pairCandidates[i * 2]
          const b = pairCandidates[i * 2 + 1]
          const first = (a.el.height || 0) <= (b.el.height || 0) ? a : b
          const second = first === a ? b : a
          setTextType(first.el, 'itemTitle')
          setTextType(second.el, 'item')
        }

        // 保底：如果没有 item，则不强行把剩余文本都标为 item（避免误用装饰文字）
      }
    }

    const getSlideTextStats = (slide: Slide) => {
      const candidates = getTextCandidates(slide).map(el => {
        const html = getHtmlText(el)
        const plain = htmlToPlainText(html)
        const font = getMaxFontSizePx(html)
        const w = (el as any).width || 0
        const h = (el as any).height || 0
        const area = w * h
        return { el, plain, font, w, h, area }
      })
      const usable = candidates
        .filter(c => !!c.plain)
        .filter(c => !isFooterLikeEl(c.el))
        .filter(c => !isBrandLikeEl(c.plain, c.el))
        .filter(c => !isTooTinyEl(c.el, c.font))

      const lowerAll = usable.map(c => (c.plain || '').toLowerCase()).join(' | ')

      const maxFont = usable.reduce((m, c) => Math.max(m, c.font), 0)
      const maxFontEl = usable.slice().sort((a, b) => b.font - a.font)[0]?.el
      const maxFontTop = maxFontEl ? ((maxFontEl as any).top || 0) : 0

      const listLike = usable.filter(c => {
        if (c.w < viewportSize.value * 0.28) return false
        if (c.h > slideHeight * 0.16) return false
        if (c.area > viewportSize.value * slideHeight * 0.25) return false
        return true
      })

      return {
        usable,
        lowerAll,
        maxFont,
        maxFontTop,
        listLikeCount: listLike.length,
      }
    }

    const scoreSlideAs = (slide: Slide, type: SlideType) => {
      const s = getSlideTextStats(slide)
      const textCount = s.usable.length
      const has = (kw: string[]) => kw.some(k => s.lowerAll.includes(k.toLowerCase()))

      if (type === 'end') {
        let score = 0
        if (has(['致谢', '谢谢', 'thanks', 'thank', 'the end', '结束'])) score += 120
        if (textCount <= 3) score += 20
        if (s.maxFont >= 36) score += 10
        return score
      }
      if (type === 'contents') {
        let score = 0
        if (has(['目录', 'contents', 'outline', 'agenda'])) score += 100
        score += Math.min(60, s.listLikeCount * 10)
        if (textCount >= 5) score += 10
        if (s.maxFont >= 28 && s.maxFontTop < slideHeight * 0.35) score += 10
        return score
      }
      if (type === 'transition') {
        let score = 0
        if (has(['章节', '部分', 'part', 'chapter', 'section'])) score += 80
        if (textCount <= 3) score += 30
        if (s.maxFont >= 40 && s.maxFontTop >= slideHeight * 0.2 && s.maxFontTop <= slideHeight * 0.7) score += 40
        if (s.listLikeCount >= 4) score -= 40
        return score
      }
      if (type === 'cover') {
        let score = 0
        if (has(['封面', 'title', '主题'])) score += 40
        if (s.maxFont >= 40 && s.maxFontTop < slideHeight * 0.35) score += 60
        if (textCount <= 4) score += 20
        if (s.listLikeCount >= 4) score -= 60
        return score
      }
      // content
      let score = 0
      if (s.listLikeCount >= 4) score += 20
      if (textCount >= 2) score += 10
      return score
    }

    const cloneSlide = (source: Slide): Slide => {
      // 深拷贝，确保不同类型模板不会互相影响
      return JSON.parse(JSON.stringify(source)) as Slide
    }

    const toTemplateSlide = (source: Slide, slideType: SlideType): Slide => {
      const s = cloneSlide(source)
      s.type = slideType
      autoTagSlideTextTypes(s, slideType)
      ensurePlaceholders(s, slideType)
      return s
    }

    const sourceSlides = slides.length ? slides : [{ id: nanoid(10), elements: [] } as Slide]

    const used = new Set<number>()
    const pickBestIndex = (type: SlideType, fallbackIndex: number, threshold: number) => {
      let bestIdx = -1
      let bestScore = -Infinity
      for (let i = 0; i < sourceSlides.length; i++) {
        if (used.has(i)) continue
        const sc = scoreSlideAs(sourceSlides[i], type)
        if (sc > bestScore) {
          bestScore = sc
          bestIdx = i
        }
      }
      if (bestIdx !== -1 && bestScore >= threshold) {
        used.add(bestIdx)
        return bestIdx
      }
      const fb = Math.min(Math.max(0, fallbackIndex), sourceSlides.length - 1)
      if (!used.has(fb)) {
        used.add(fb)
        return fb
      }
      // fallback 冲突：找一个没用过的
      for (let i = 0; i < sourceSlides.length; i++) {
        if (!used.has(i)) {
          used.add(i)
          return i
        }
      }
      return fb
    }

    // 先挑 end，尽量用“谢谢/致谢”页；否则再兜底最后一页
    const endIndex = pickBestIndex('end', sourceSlides.length - 1, 80)
    const coverIndex = pickBestIndex('cover', 0, 60)
    const contentsIndex = pickBestIndex('contents', 1, 60)
    const transitionIndex = pickBestIndex('transition', 2, 50)

    const coverSlide = toTemplateSlide(sourceSlides[coverIndex], 'cover')
    const contentsSlide = toTemplateSlide(sourceSlides[contentsIndex], 'contents')
    const transitionSlide = toTemplateSlide(sourceSlides[transitionIndex], 'transition')
    const endSlide = toTemplateSlide(sourceSlides[endIndex], 'end')

    let contentSources = sourceSlides
      .map((s, i) => ({ s, i }))
      .filter(({ i }) => ![coverIndex, contentsIndex, transitionIndex, endIndex].includes(i))
      .map(({ s }) => s)

    if (!contentSources.length) contentSources = [sourceSlides[coverIndex]]
    const contentSlides = contentSources.map(s => toTemplateSlide(s, 'content'))

    const templateSlides = [coverSlide, contentsSlide, transitionSlide, ...contentSlides, endSlide]

    return {
      slides: templateSlides,
      theme: templateTheme,
    }
  }

  // 导入JSON文件
  const importJSON = (files: FileList | File[], cover = false) => {
    const file = files[0]

    const reader = new FileReader()
    reader.addEventListener('load', () => {
      try {
        const { slides, theme } = JSON.parse(reader.result as string)
        if (cover) {
          slidesStore.updateSlideIndex(0)
          slidesStore.setSlides(slides, (theme || {}))
          addHistorySnapshot()
        }
        else if (isEmptySlide.value) {
          slidesStore.setSlides(slides, (theme || {}))
          addHistorySnapshot()
        }
        else addSlidesFromData(slides)
      }
      catch {
        message.error('无法正确读取 / 解析该文件')
      }
    })
    reader.readAsText(file)
  }

  // 导入pptist文件
  const importSpecificFile = (files: FileList | File[], cover = false) => {
    const file = files[0]

    const reader = new FileReader()
    reader.addEventListener('load', () => {
      try {
        const { slides, theme } = JSON.parse(decrypt(reader.result as string))
        if (cover) {
          slidesStore.updateSlideIndex(0)
          slidesStore.setSlides(slides, (theme || {}))
          addHistorySnapshot()
        }
        else if (isEmptySlide.value) {
          slidesStore.setSlides(slides, (theme || {}))
          addHistorySnapshot()
        }
        else addSlidesFromData(slides)
      }
      catch {
        message.error('无法正确读取 / 解析该文件')
      }
    })
    reader.readAsText(file)
  }

  const rotateLine = (line: PPTLineElement, angleDeg: number) => {
    const { start, end } = line
      
    const angleRad = angleDeg * Math.PI / 180
    
    const midX = (start[0] + end[0]) / 2
    const midY = (start[1] + end[1]) / 2
    
    const startTransX = start[0] - midX
    const startTransY = start[1] - midY
    const endTransX = end[0] - midX
    const endTransY = end[1] - midY
    
    const cosA = Math.cos(angleRad)
    const sinA = Math.sin(angleRad)
    
    const startRotX = startTransX * cosA - startTransY * sinA
    const startRotY = startTransX * sinA + startTransY * cosA
    
    const endRotX = endTransX * cosA - endTransY * sinA
    const endRotY = endTransX * sinA + endTransY * cosA
    
    const startNewX = startRotX + midX
    const startNewY = startRotY + midY
    const endNewX = endRotX + midX
    const endNewY = endRotY + midY
    
    const beforeMinX = Math.min(start[0], end[0])
    const beforeMinY = Math.min(start[1], end[1])
    
    const afterMinX = Math.min(startNewX, endNewX)
    const afterMinY = Math.min(startNewY, endNewY)
    
    const startAdjustedX = startNewX - afterMinX
    const startAdjustedY = startNewY - afterMinY
    const endAdjustedX = endNewX - afterMinX
    const endAdjustedY = endNewY - afterMinY
    
    const startAdjusted: [number, number] = [startAdjustedX, startAdjustedY]
    const endAdjusted: [number, number] = [endAdjustedX, endAdjustedY]
    const offset = [afterMinX - beforeMinX, afterMinY - beforeMinY]
    
    return {
      start: startAdjusted,
      end: endAdjusted,
      offset,
    }
  }

  const parseLineElement = (el: Shape, ratio: number) => {
    let start: [number, number] = [0, 0]
    let end: [number, number] = [0, 0]

    if (!el.isFlipV && !el.isFlipH) { // 右下
      start = [0, 0]
      end = [el.width, el.height]
    }
    else if (el.isFlipV && el.isFlipH) { // 左上
      start = [el.width, el.height]
      end = [0, 0]
    }
    else if (el.isFlipV && !el.isFlipH) { // 右上
      start = [0, el.height]
      end = [el.width, 0]
    }
    else { // 左下
      start = [el.width, 0]
      end = [0, el.height]
    }

    const data: PPTLineElement = {
      type: 'line',
      id: nanoid(10),
      width: +((el.borderWidth || 1) * ratio).toFixed(2),
      left: el.left,
      top: el.top,
      start,
      end,
      style: el.borderType,
      color: el.borderColor,
      points: ['', /straightConnector/.test(el.shapType) ? 'arrow' : '']
    }
    if (el.rotate) {
      const { start, end, offset } = rotateLine(data, el.rotate)

      data.start = start
      data.end = end
      data.left = data.left + offset[0]
      data.top = data.top + offset[1]
    }
    if (/bentConnector/.test(el.shapType)) {
      data.broken2 = [
        Math.abs(data.start[0] - data.end[0]) / 2,
        Math.abs(data.start[1] - data.end[1]) / 2,
      ]
    }
    if (/curvedConnector/.test(el.shapType)) {
      const cubic: [number, number] = [
        Math.abs(data.start[0] - data.end[0]) / 2,
        Math.abs(data.start[1] - data.end[1]) / 2,
      ]
      data.cubic = [cubic, cubic]
    }

    return data
  }

  const flipGroupElements = (elements: BaseElement[], axis: 'x' | 'y') => {
    const minX = Math.min(...elements.map(el => el.left))
    const maxX = Math.max(...elements.map(el => el.left + el.width))
    const minY = Math.min(...elements.map(el => el.top))
    const maxY = Math.max(...elements.map(el => el.top + el.height))

    const centerX = (minX + maxX) / 2
    const centerY = (minY + maxY) / 2

    return elements.map(element => {
      const newElement = { ...element }

      if (axis === 'y') newElement.left = 2 * centerX - element.left - element.width
      if (axis === 'x') newElement.top = 2 * centerY - element.top - element.height
  
      return newElement
    })
  }

  const calculateRotatedPosition = (
    x: number,
    y: number,
    w: number,
    h: number,
    ox: number,
    oy: number,
    k: number,
  ) => {
    const radians = k * (Math.PI / 180)

    const containerCenterX = x + w / 2
    const containerCenterY = y + h / 2

    const relativeX = ox - w / 2
    const relativeY = oy - h / 2

    const rotatedX = relativeX * Math.cos(radians) + relativeY * Math.sin(radians)
    const rotatedY = -relativeX * Math.sin(radians) + relativeY * Math.cos(radians)

    const graphicX = containerCenterX + rotatedX
    const graphicY = containerCenterY + rotatedY

    return { x: graphicX, y: graphicY }
  }

  // 导入PPTX文件
  const importPPTXFile = (files: FileList | File[], options?: { cover?: boolean; fixedViewport?: boolean }) => {
    const defaultOptions = {
      cover: false,
      fixedViewport: false, 
    }
    const { cover, fixedViewport } = { ...defaultOptions, ...options }

    const file = files[0]
    if (!file) return

    exporting.value = true

    const shapeList: ShapePoolItem[] = []
    for (const item of SHAPE_LIST) {
      shapeList.push(...item.children)
    }
    
    const reader = new FileReader()
    reader.onload = async e => {
      let json = null
      try {
        json = await parse(e.target!.result as ArrayBuffer)
      }
      catch {
        exporting.value = false
        message.error('无法正确读取 / 解析该文件')
        return
      }

      let ratio = 96 / 72
      const width = json.size.width
      
      if (fixedViewport) ratio = 1000 / width
      else slidesStore.setViewportSize(width * ratio)

      slidesStore.setTheme({ themeColors: json.themeColors })

      const slides: Slide[] = []
      for (const item of json.slides) {
        const { type, value } = item.fill
        let background: SlideBackground
        if (type === 'image') {
          background = {
            type: 'image',
            image: {
              src: value.picBase64,
              size: 'cover',
            },
          }
        }
        else if (type === 'gradient') {
          background = {
            type: 'gradient',
            gradient: {
              type: value.path === 'line' ? 'linear' : 'radial',
              colors: value.colors.map(item => ({
                ...item,
                pos: parseInt(item.pos),
              })),
              rotate: value.rot + 90,
            },
          }
        }
        else if (type === 'pattern') {
          background = {
            type: 'solid',
            color: '#fff',
          }
        }
        else {
          background = {
            type: 'solid',
            color: value || '#fff',
          }
        }

        const slide: Slide = {
          id: nanoid(10),
          elements: [],
          background,
          remark: item.note || '',
        }

        const parseElements = (elements: Element[]) => {
          const sortedElements = elements.sort((a, b) => a.order - b.order)

          for (const el of sortedElements) {
            const originWidth = el.width || 1
            const originHeight = el.height || 1
            const originLeft = el.left
            const originTop = el.top

            el.width = el.width * ratio
            el.height = el.height * ratio
            el.left = el.left * ratio
            el.top = el.top * ratio
  
            if (el.type === 'text') {
              if (el.autoFit && el.autoFit.type === 'text') {
                const fontScale = ratio * (el.autoFit.fontScale || 100) / 100
                const shapeEl: PPTShapeElement = {
                  type: 'shape',
                  id: nanoid(10),
                  width: el.width,
                  height: el.height,
                  left: el.left,
                  top: el.top,
                  rotate: el.rotate,
                  viewBox: [200, 200],
                  path: 'M 0 0 L 200 0 L 200 200 L 0 200 Z',
                  fill: el.fill.type === 'color' ? el.fill.value : '',
                  fixedRatio: false,
                  outline: {
                    color: el.borderColor,
                    width: +(el.borderWidth * ratio).toFixed(2),
                    style: el.borderType,
                  },
                  text: {
                    content: convertFontSizePtToPx(el.content, fontScale),
                    defaultFontName: theme.value.fontName,
                    defaultColor: theme.value.fontColor,
                    align: shapeVAlignMap[el.vAlign] || 'middle',
                    lineHeight: 1,
                  },
                }
                slide.elements.push(shapeEl)
              }
              else {
                const textEl: PPTTextElement = {
                  type: 'text',
                  id: nanoid(10),
                  width: el.width,
                  height: el.height,
                  left: el.left,
                  top: el.top,
                  rotate: el.rotate,
                  defaultFontName: theme.value.fontName,
                  defaultColor: theme.value.fontColor,
                  content: convertFontSizePtToPx(el.content, ratio),
                  lineHeight: 1,
                  outline: {
                    color: el.borderColor,
                    width: +(el.borderWidth * ratio).toFixed(2),
                    style: el.borderType,
                  },
                  fill: el.fill.type === 'color' ? el.fill.value : '',
                  vertical: el.isVertical,
                }
                if (el.shadow) {
                  textEl.shadow = {
                    h: el.shadow.h * ratio,
                    v: el.shadow.v * ratio,
                    blur: el.shadow.blur * ratio,
                    color: el.shadow.color,
                  }
                }
                slide.elements.push(textEl)
              }
            }
            else if (el.type === 'image') {
              const element: PPTImageElement = {
                type: 'image',
                id: nanoid(10),
                src: el.src,
                width: el.width,
                height: el.height,
                left: el.left,
                top: el.top,
                fixedRatio: true,
                rotate: el.rotate,
                flipH: el.isFlipH,
                flipV: el.isFlipV,
              }
              if (el.borderWidth) {
                element.outline = {
                  color: el.borderColor,
                  width: +(el.borderWidth * ratio).toFixed(2),
                  style: el.borderType,
                }
              }
              const clipShapeTypes = ['roundRect', 'ellipse', 'triangle', 'rhombus', 'pentagon', 'hexagon', 'heptagon', 'octagon', 'parallelogram', 'trapezoid']
              if (el.rect) {
                element.clip = {
                  shape: (el.geom && clipShapeTypes.includes(el.geom)) ? el.geom : 'rect',
                  range: [
                    [
                      el.rect.l || 0,
                      el.rect.t || 0,
                    ],
                    [
                      100 - (el.rect.r || 0),
                      100 - (el.rect.b || 0),
                    ],
                  ]
                }
              }
              else if (el.geom && clipShapeTypes.includes(el.geom)) {
                element.clip = {
                  shape: el.geom,
                  range: [[0, 0], [100, 100]]
                }
              }
              slide.elements.push(element)
            }
            else if (el.type === 'math') {
              slide.elements.push({
                type: 'image',
                id: nanoid(10),
                src: el.picBase64,
                width: el.width,
                height: el.height,
                left: el.left,
                top: el.top,
                fixedRatio: true,
                rotate: 0,
              })
            }
            else if (el.type === 'audio') {
              slide.elements.push({
                type: 'audio',
                id: nanoid(10),
                src: el.blob,
                width: el.width,
                height: el.height,
                left: el.left,
                top: el.top,
                rotate: 0,
                fixedRatio: false,
                color: theme.value.themeColors[0],
                loop: false,
                autoplay: false,
              })
            }
            else if (el.type === 'video') {
              slide.elements.push({
                type: 'video',
                id: nanoid(10),
                src: (el.blob || el.src)!,
                width: el.width,
                height: el.height,
                left: el.left,
                top: el.top,
                rotate: 0,
                autoplay: false,
              })
            }
            else if (el.type === 'shape') {
              if (el.shapType === 'line' || /Connector/.test(el.shapType)) {
                const lineElement = parseLineElement(el, ratio)
                slide.elements.push(lineElement)
              }
              else {
                const shape = shapeList.find(item => item.pptxShapeType === el.shapType)

                const gradient: Gradient | undefined = el.fill?.type === 'gradient' ? {
                  type: el.fill.value.path === 'line' ? 'linear' : 'radial',
                  colors: el.fill.value.colors.map(item => ({
                    ...item,
                    pos: parseInt(item.pos),
                  })),
                  rotate: el.fill.value.rot,
                } : undefined

                const pattern: string | undefined = el.fill?.type === 'image' ? el.fill.value.picBase64 : undefined

                const fill = el.fill?.type === 'color' ? el.fill.value : ''
                
                const element: PPTShapeElement = {
                  type: 'shape',
                  id: nanoid(10),
                  width: el.width,
                  height: el.height,
                  left: el.left,
                  top: el.top,
                  viewBox: [200, 200],
                  path: 'M 0 0 L 200 0 L 200 200 L 0 200 Z',
                  fill,
                  gradient,
                  pattern,
                  fixedRatio: false,
                  rotate: el.rotate,
                  outline: {
                    color: el.borderColor,
                    width: +(el.borderWidth * ratio).toFixed(2),
                    style: el.borderType,
                  },
                  text: {
                    content: convertFontSizePtToPx(el.content, ratio),
                    defaultFontName: theme.value.fontName,
                    defaultColor: theme.value.fontColor,
                    align: shapeVAlignMap[el.vAlign] || 'middle',
                  },
                  flipH: el.isFlipH,
                  flipV: el.isFlipV,
                }
                if (el.shadow) {
                  element.shadow = {
                    h: el.shadow.h * ratio,
                    v: el.shadow.v * ratio,
                    blur: el.shadow.blur * ratio,
                    color: el.shadow.color,
                  }
                }
    
                if (shape) {
                  element.path = shape.path
                  element.viewBox = shape.viewBox
    
                  if (shape.pathFormula) {
                    element.pathFormula = shape.pathFormula
                    element.viewBox = [el.width, el.height]
    
                    const pathFormula = SHAPE_PATH_FORMULAS[shape.pathFormula]
                    if ('editable' in pathFormula && pathFormula.editable) {
                      element.path = pathFormula.formula(el.width, el.height, pathFormula.defaultValue)
                      element.keypoints = pathFormula.defaultValue
                    }
                    else element.path = pathFormula.formula(el.width, el.height)
                  }
                }
                else if (el.path && el.path.indexOf('NaN') === -1) {
                  const { maxX, maxY } = getSvgPathRange(el.path)
                  element.path = el.path
                  if ((maxX / maxY) > (originWidth / originHeight)) {
                    element.viewBox = [maxX, maxX * originHeight / originWidth]
                  }
                  else {
                    element.viewBox = [maxY * originWidth / originHeight, maxY]
                  }
                }
                if (el.shapType === 'custom') {
                  if (el.path!.indexOf('NaN') !== -1) {
                    if (element.width === 0) element.width = 0.1
                    if (element.height === 0) element.height = 0.1
                    element.path = el.path!.replace(/NaN/g, '0')
                  }
                  else {
                    element.special = true
                    element.path = el.path!
                  }
                  const { maxX, maxY } = getSvgPathRange(element.path)
                  if ((maxX / maxY) > (originWidth / originHeight)) {
                    element.viewBox = [maxX, maxX * originHeight / originWidth]
                  }
                  else {
                    element.viewBox = [maxY * originWidth / originHeight, maxY]
                  }
                }
    
                if (element.path) slide.elements.push(element)
              }
            }
            else if (el.type === 'table') {
              const row = el.data.length
              const col = el.data[0].length
  
              const style: TableCellStyle = {
                fontname: theme.value.fontName,
                color: theme.value.fontColor,
              }
              const data: TableCell[][] = []
              for (let i = 0; i < row; i++) {
                const rowCells: TableCell[] = []
                for (let j = 0; j < col; j++) {
                  const cellData = el.data[i][j]

                  let textDiv: HTMLDivElement | null = document.createElement('div')
                  textDiv.innerHTML = cellData.text
                  const p = textDiv.querySelector('p')
                  const align = p?.style.textAlign || 'left'

                  const span = textDiv.querySelector('span')
                  const fontsize = span?.style.fontSize ? (parseInt(span?.style.fontSize) * ratio).toFixed(1) + 'px' : ''
                  const fontname = span?.style.fontFamily || ''
                  const color = span?.style.color || cellData.fontColor

                  rowCells.push({
                    id: nanoid(10),
                    colspan: cellData.colSpan || 1,
                    rowspan: cellData.rowSpan || 1,
                    text: textDiv.innerText,
                    style: {
                      ...style,
                      align: ['left', 'right', 'center'].includes(align) ? (align as 'left' | 'right' | 'center') : 'left',
                      fontsize,
                      fontname,
                      color,
                      bold: cellData.fontBold,
                      backcolor: cellData.fillColor,
                    },
                  })
                  textDiv = null
                }
                data.push(rowCells)
              }
  
              const allWidth = el.colWidths.reduce((a, b) => a + b, 0)
              const colWidths: number[] = el.colWidths.map(item => item / allWidth)

              const firstCell = el.data[0][0]
              const border = firstCell.borders.top ||
                firstCell.borders.bottom ||
                el.borders.top ||
                el.borders.bottom ||
                firstCell.borders.left ||
                firstCell.borders.right ||
                el.borders.left ||
                el.borders.right
              const borderWidth = border?.borderWidth || 0
              const borderStyle = border?.borderType || 'solid'
              const borderColor = border?.borderColor || '#eeece1'
  
              slide.elements.push({
                type: 'table',
                id: nanoid(10),
                width: el.width,
                height: el.height,
                left: el.left,
                top: el.top,
                colWidths,
                rotate: 0,
                data,
                outline: {
                  width: +(borderWidth * ratio || 2).toFixed(2),
                  style: borderStyle,
                  color: borderColor,
                },
                cellMinHeight: el.rowHeights[0] ? el.rowHeights[0] * ratio : 36,
              })
            }
            else if (el.type === 'chart') {
              let labels: string[]
              let legends: string[]
              let series: number[][]
  
              if (el.chartType === 'scatterChart' || el.chartType === 'bubbleChart') {
                labels = el.data[0].map((item, index) => `坐标${index + 1}`)
                legends = ['X', 'Y']
                series = el.data
              }
              else {
                const data = el.data as ChartItem[]
                labels = Object.values(data[0].xlabels)
                legends = data.map(item => item.key)
                series = data.map(item => item.values.map(v => v.y))
              }

              const options: ChartOptions = {}
  
              let chartType: ChartType = 'bar'

              switch (el.chartType) {
                case 'barChart':
                case 'bar3DChart':
                  chartType = 'bar'
                  if (el.barDir === 'bar') chartType = 'column'
                  if (el.grouping === 'stacked' || el.grouping === 'percentStacked') options.stack = true
                  break
                case 'lineChart':
                case 'line3DChart':
                  if (el.grouping === 'stacked' || el.grouping === 'percentStacked') options.stack = true
                  chartType = 'line'
                  break
                case 'areaChart':
                case 'area3DChart':
                  if (el.grouping === 'stacked' || el.grouping === 'percentStacked') options.stack = true
                  chartType = 'area'
                  break
                case 'scatterChart':
                case 'bubbleChart':
                  chartType = 'scatter'
                  break
                case 'pieChart':
                case 'pie3DChart':
                  chartType = 'pie'
                  break
                case 'radarChart':
                  chartType = 'radar'
                  break
                case 'doughnutChart':
                  chartType = 'ring'
                  break
                default:
              }
  
              slide.elements.push({
                type: 'chart',
                id: nanoid(10),
                chartType: chartType,
                width: el.width,
                height: el.height,
                left: el.left,
                top: el.top,
                rotate: 0,
                themeColors: el.colors.length ? el.colors : theme.value.themeColors,
                textColor: theme.value.fontColor,
                data: {
                  labels,
                  legends,
                  series,
                },
                options,
              })
            }
            else if (el.type === 'group') {
              let elements: BaseElement[] = el.elements.map(_el => {
                let left = _el.left + originLeft
                let top = _el.top + originTop

                if (el.rotate) {
                  const { x, y } = calculateRotatedPosition(originLeft, originTop, originWidth, originHeight, _el.left, _el.top, el.rotate)
                  left = x
                  top = y
                }

                const element = {
                  ..._el,
                  left,
                  top,
                }
                if (el.isFlipH && 'isFlipH' in element) element.isFlipH = true
                if (el.isFlipV && 'isFlipV' in element) element.isFlipV = true

                return element
              })
              if (el.isFlipH) elements = flipGroupElements(elements, 'y')
              if (el.isFlipV) elements = flipGroupElements(elements, 'x')
              parseElements(elements)
            }
            else if (el.type === 'diagram') {
              const elements = el.elements.map(_el => ({
                ..._el,
                left: _el.left + originLeft,
                top: _el.top + originTop,
              }))
              parseElements(elements)
            }
          }
        }
        parseElements([...item.elements, ...item.layoutElements])
        slides.push(slide)
      }

      if (cover) {
        slidesStore.updateSlideIndex(0)
        slidesStore.setSlides(slides)
        addHistorySnapshot()
      }
      else if (isEmptySlide.value) {
        slidesStore.setSlides(slides)
        addHistorySnapshot()
      }
      else addSlidesFromData(slides)

      exporting.value = false
    }
    reader.readAsArrayBuffer(file)
  }

  return {
    importSpecificFile,
    importJSON,
    importPPTXFile,
    parsePPTXTemplate,
    exporting,
  }
}