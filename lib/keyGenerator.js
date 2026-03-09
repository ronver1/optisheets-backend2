'use strict';

const crypto = require('crypto');

const ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';

function randomSegment(length) {
  let result = '';
  // Generate more bytes than needed to avoid modulo bias
  const bytes = crypto.randomBytes(length * 4);
  let used = 0;
  while (result.length < length) {
    const byte = bytes[used++];
    // Reject values that would introduce bias (256 % 36 = 4, discard >= 252)
    if (byte < 252) {
      result += ALPHABET[byte % ALPHABET.length];
    }
  }
  return result;
}

function generatePrivateKey(templateId) {
  const templateCode = templateId.replace(/[^a-zA-Z]/g, '').slice(0, 3).toUpperCase();
  const seg1 = randomSegment(4);
  const seg2 = randomSegment(4);
  const seg3 = randomSegment(4);
  return `OS-${templateCode}-${seg1}-${seg2}-${seg3}`;
}

const KEY_PATTERN = /^OS-[A-Z]{3}-[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4}$/;

function validateKeyFormat(key) {
  return KEY_PATTERN.test(key);
}

module.exports = { generatePrivateKey, validateKeyFormat };
