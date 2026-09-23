// Common slide intermediate representation (IR).
function createDeck(meta = {}) {
  return { version: 1, meta, slides: [] };
}
function createSlide({ id, width = 1920, height = 1080, elements = [] } = {}) {
  return { id: id || `slide-${Date.now()}`, width, height, elements };
}
function normalizeElement(el) {
  return {
    id: el.id || null,
    type: el.type || 'unknown',
    x: Number(el.x || 0), y: Number(el.y || 0),
    width: Number(el.width || 0), height: Number(el.height || 0),
    text: el.text || '',
    style: el.style || {},
    source: el.source || null,
  };
}
function normalizeDeck(deck) {
  return {
    version: 1,
    meta: deck.meta || {},
    slides: (deck.slides || []).map(s => ({
      ...createSlide(s),
      elements: (s.elements || []).map(normalizeElement),
    })),
  };
}
module.exports = { createDeck, createSlide, normalizeElement, normalizeDeck };
