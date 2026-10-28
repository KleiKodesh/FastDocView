// The demo's sample files, kept in the repository's samples/ folder (paths relative to the repo
// root; the dev server resolves them). The page opens the .doc on first load; a later pick is
// remembered in localStorage and wins.
export const DEFAULT_FOLDER_PATH = 'samples'

export interface WordDemoSample { label: string; path: string }

export const SAMPLE_FILES: WordDemoSample[] = [
  { label: '.doc — 40 MB, 107k paragraphs', path: 'samples/dor-hamelaktim-shabbat.doc' },
  { label: '.docx — 87 MB body, 123k blocks', path: 'samples/dor-hamelaktim-shabbat.docx' },
]

export const DEFAULT_FILE_PATH = SAMPLE_FILES[0]!.path
export const LAST_FILE_STORAGE_KEY = 'word-demo.file'
