'use strict';

const TEMPLATES = {
  'internship-tracker': {
    id: 'internship-tracker',
    name: 'Internship Tracker',
    isAI: true,
    sheetUrl: 'https://docs.google.com/spreadsheets/d/1cXhCNOJxj5DE-P_h2gE1zD_jyMpwwWW9vea79ZCdX0w/edit?usp=sharing',
    pdfUrl: 'https://docs.google.com/document/d/14oo5npJCbEuK8X3yP358kHGhW52B7dFK/edit?usp=sharing&ouid=106805336580345169290&rtpof=true&sd=true',
    etsy_title_keywords: ['internship tracker', 'internship-tracker'],
  },
  'internship-tracker-standard': {
    id: 'internship-tracker-standard',
    name: 'Internship Tracker (Standard)',
    isAI: false,
    sheetUrl: 'https://placeholder.example.com/sheets/internship-tracker-standard',
    pdfUrl: 'https://placeholder.example.com/pdf/internship-tracker-standard',
    etsy_title_keywords: ['internship tracker standard', 'internship-tracker-standard'],
  },
};

const CREDIT_PACKS = {
  5: {
    amount: 5,
    etsy_title_keywords: ['5 credits', '5-credit', '5 pack'],
  },
  10: {
    amount: 10,
    etsy_title_keywords: ['10 credits', '10-credit', '10 pack'],
  },
  25: {
    amount: 25,
    etsy_title_keywords: ['25 credits', '25-credit', '25 pack'],
  },
  50: {
    amount: 50,
    etsy_title_keywords: ['50 credits', '50-credit', '50 pack'],
  },
};

module.exports = { TEMPLATES, CREDIT_PACKS };
