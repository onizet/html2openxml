import { defaultTheme } from '@vuepress/theme-default'
import { defineUserConfig } from 'vuepress'
import { viteBundler } from '@vuepress/bundler-vite'
import { llmsPlugin } from '@vuepress/plugin-llms'

export default defineUserConfig({
  lang: 'en-US',

  title: 'Html to OpenXml',
  description: 'Faithfully translates rich HTML markup into professional OpenXML DOCX files.',

  base: '/html2openxml/',
  theme: defaultTheme({
    editLink: true,
    docsRepo: 'https://github.com/onizet/html2openxml',
    docsBranch: 'master',
    docsDir: 'docs',
    logo: 'images/hero.png',

    navbar: ['/', 
      { text: 'API Reference', link: '/guide/api' },
      { text: 'LLM Docs Friendly', link: '/llmdoc' },
      { text: 'NuGet', link: 'https://www.nuget.org/packages/HtmlToOpenXml', ariaLabel: 'NuGet', target: '_blank' }
    ],

    sidebar: [
      {
        text: 'API Reference',
        prefix: 'guide/',
        children: ['quickstart', 'api', 'hyperlinks', 'images', 'styling', 'tables', 'numbering', 'pre', 'footnotes', 'pagebreak', 'rtl'],
      },
      {
        text: 'LLM Docs Friendly', link: 'llmdoc'
      },
      {
        text: 'Performance', link: 'appendix/performance'
      },
      {
        text: 'Appendix',
        collapsible: true,
        prefix: 'appendix/',
        children: ['content-type', 'template-to-docx', 'placeholder-replacement', 'restrict-edition'],
      },
    ]
  }),

  bundler: viteBundler(),

  plugins: [
    llmsPlugin({
    }),
  ]
})
