// Windowed renderer for the test viewer. Every block has an estimated height (from its text
// length) until it has been rendered once and measured; a prefix-sum table over those heights
// gives block ↔ pixel mapping. Only [firstVisible - OVERSCAN, lastVisible + OVERSCAN] exists in
// the DOM; spacers above and below hold the rest of the scroll height. Requests to the host are
// coalesced: at most one in flight, the newest wins.
//
// This is a stand-in for the app's Vue windowed scroller (see NOTES.md, "Integration").

const OVERSCAN = 15;
const AVERAGE_CHARACTERS_PER_LINE = 85;
const LINE_HEIGHT_PX = 22;
const PARAGRAPH_PADDING_PX = 8;

const scroller = document.getElementById('scroller');
const topSpacer = document.getElementById('top-spacer');
const bottomSpacer = document.getElementById('bottom-spacer');
const windowElement = document.getElementById('window');
const positionLabel = document.getElementById('position');
const gotoInput = document.getElementById('goto-block');
const headingsList = document.getElementById('headings-list');
const fileNameLabel = document.getElementById('file-name');
const documentStylesheet = document.getElementById('document-stylesheet');

let blockCount = 0;
let heights = new Float64Array(0);          // estimated or measured height per block
let prefixHeights = new Float64Array(0);    // prefixHeights[i] = top of block i; length blockCount + 1
let isPrefixDirty = false;
let measured = new Uint8Array(0);
let renderedFirst = 0, renderedCount = 0;
let htmlCache = new Map();                  // index → html, bounded
let inFlightRequestId = 0;
let pendingRange = null;

function post(message) { window.chrome.webview.postMessage(message); }

window.chrome.webview.addEventListener('message', event => {
  const message = event.data;
  if (message.type === 'document') onDocument(message);
  else if (message.type === 'blocks') onBlocks(message);
  else if (message.type === 'closed') onClosed();
});

function onDocument(message) {
  fileNameLabel.textContent = message.fileName;
  documentStylesheet.textContent = message.stylesheet;
  windowElement.dir = message.isRightToLeft ? 'rtl' : 'ltr';
  blockCount = message.blockCount;
  heights = new Float64Array(blockCount);
  measured = new Uint8Array(blockCount);
  for (let i = 0; i < blockCount; i++) heights[i] = estimateHeight(message.estimatedTextLengths[i]);
  prefixHeights = new Float64Array(blockCount + 1);
  rebuildPrefix();
  htmlCache = new Map();
  renderedFirst = 0; renderedCount = 0;
  windowElement.innerHTML = '';
  gotoInput.max = Math.max(0, blockCount - 1);
  renderHeadings(message.headings);
  scroller.scrollTop = 0;
  requestVisibleWindow();
}

function onClosed() {
  blockCount = 0;
  windowElement.innerHTML = '';
  topSpacer.style.height = bottomSpacer.style.height = '0px';
  headingsList.innerHTML = '';
  fileNameLabel.textContent = 'No document';
}

function estimateHeight(textLength) {
  return PARAGRAPH_PADDING_PX + Math.max(1, Math.ceil(textLength / AVERAGE_CHARACTERS_PER_LINE)) * LINE_HEIGHT_PX;
}

function rebuildPrefix() {
  let total = 0;
  for (let i = 0; i < blockCount; i++) { prefixHeights[i] = total; total += heights[i]; }
  prefixHeights[blockCount] = total;
  isPrefixDirty = false;
}

function blockAtOffset(offset) {
  let low = 0, high = blockCount - 1;
  while (low < high) {
    const middle = (low + high + 1) >> 1;
    if (prefixHeights[middle] <= offset) low = middle; else high = middle - 1;
  }
  return low;
}

function requestVisibleWindow() {
  if (blockCount === 0) return;
  if (isPrefixDirty) rebuildPrefix();
  const toolbarHeight = document.getElementById('toolbar').offsetHeight;
  const top = Math.max(0, scroller.scrollTop - toolbarHeight);
  const bottom = top + scroller.clientHeight;
  const first = Math.max(0, blockAtOffset(top) - OVERSCAN);
  const last = Math.min(blockCount - 1, blockAtOffset(bottom) + OVERSCAN);
  positionLabel.textContent = `block ${blockAtOffset(top).toLocaleString()} of ${blockCount.toLocaleString()}`;
  if (first >= renderedFirst && last < renderedFirst + renderedCount && renderedCount > 0) return;   // already in the DOM

  pendingRange = { first, count: last - first + 1 };
  if (inFlightRequestId === 0) sendPendingRequest();
}

function sendPendingRequest() {
  if (!pendingRange) return;
  const range = pendingRange;
  pendingRange = null;
  // blocks already cached need no round trip, but keep it simple: ask for the whole range
  inFlightRequestId = Date.now() % 1_000_000_000 + 1;
  post({ type: 'requestBlocks', first: range.first, count: range.count, requestId: inFlightRequestId });
}

function onBlocks(message) {
  if (message.requestId !== inFlightRequestId) return;   // stale
  inFlightRequestId = 0;
  renderWindow(message.first, message.blocks);
  if (pendingRange) sendPendingRequest();
  else requestVisibleWindow();                            // the measured heights may have moved the viewport
}

function renderWindow(first, blocks) {
  const anchorIndex = blockAtOffset(Math.max(0, scroller.scrollTop));
  const anchorTopBefore = prefixHeights[anchorIndex];
  const anchorScreenOffset = scroller.scrollTop - anchorTopBefore;

  const fragment = document.createDocumentFragment();
  for (const block of blocks) {
    const wrapper = document.createElement('div');
    wrapper.dataset.block = block.index;
    wrapper.innerHTML = block.html;
    fragment.appendChild(wrapper);
    if (htmlCache.size > 2000) htmlCache.clear();
    htmlCache.set(block.index, block.html);
  }
  windowElement.replaceChildren(fragment);
  renderedFirst = first;
  renderedCount = blocks.length;

  // measure and correct the estimates
  for (const wrapper of windowElement.children) {
    const index = Number(wrapper.dataset.block);
    const height = wrapper.getBoundingClientRect().height;
    if (!measured[index] || Math.abs(heights[index] - height) > 0.5) { heights[index] = height; measured[index] = 1; isPrefixDirty = true; }
  }
  if (isPrefixDirty) rebuildPrefix();
  topSpacer.style.height = prefixHeights[first] + 'px';
  bottomSpacer.style.height = (prefixHeights[blockCount] - prefixHeights[first + blocks.length]) + 'px';

  // keep the anchor block where it was on screen
  const anchorTopAfter = prefixHeights[anchorIndex];
  if (Math.abs(anchorTopAfter - anchorTopBefore) > 0.5) scroller.scrollTop = anchorTopAfter + anchorScreenOffset;
}

function scrollToBlock(index) {
  if (blockCount === 0) return;
  index = Math.max(0, Math.min(blockCount - 1, index));
  if (isPrefixDirty) rebuildPrefix();
  scroller.scrollTop = prefixHeights[index] + document.getElementById('toolbar').offsetHeight;
  requestVisibleWindow();
}

function renderHeadings(headings) {
  headingsList.innerHTML = '';
  for (const heading of headings) {
    const item = document.createElement('li');
    item.dataset.level = heading.level;
    item.textContent = heading.text;
    item.title = `block ${heading.blockIndex}`;
    item.addEventListener('click', () => scrollToBlock(heading.blockIndex));
    headingsList.appendChild(item);
  }
}

let scrollFrame = 0;
scroller.addEventListener('scroll', () => {
  if (scrollFrame) return;
  scrollFrame = requestAnimationFrame(() => { scrollFrame = 0; requestVisibleWindow(); });
});
window.addEventListener('resize', requestVisibleWindow);
gotoInput.addEventListener('change', () => scrollToBlock(Number(gotoInput.value)));

post({ type: 'ready' });
