// The demo's sample files: the largest real Word files on the dev laptop, one per format.
// The page opens the .doc on first load; a later pick is remembered in localStorage and wins.
export const DEFAULT_FOLDER_PATH = "C:\\Users\\Admin\\Documents\\גמ''ח\\ספריה תורנית\\אוצר התורה מיר - מאגר מ''מ"

export interface WordDemoSample { label: string; path: string }

export const SAMPLE_FILES: WordDemoSample[] = [
  { label: '.doc — 40 MB, 107k paragraphs', path: `${DEFAULT_FOLDER_PATH}\\3שבת\\שבת דור המלקטים.doc` },
  { label: '.docx — 87 MB body, 123k blocks', path: `${DEFAULT_FOLDER_PATH}\\0כללי\\דור המלקטים\\שבת.docx` },
]

export const DEFAULT_FILE_PATH = SAMPLE_FILES[0].path
export const LAST_FILE_STORAGE_KEY = 'word-demo.file'
