'use strict';

const TEMPLATES = {
  'internship-tracker': {
    id: 'internship-tracker',
    name: 'Internship Tracker',
    isAI: true,
    sheetUrl: 'https://docs.google.com/spreadsheets/d/127ZB9y50UFmF5RjeOddf_rKjl7TsOLsDKb3K-2ZNpWs/edit?usp=sharing',
    pdfUrl: 'https://drive.google.com/file/d/1TLMAXyR591J8xHC7KnyTgv-PrZ-SFJWM/view?usp=sharing',
    etsy_title_keywords: ['internship tracker', 'internship-tracker'],
  },
  'gpa-and-grade-tracker': {
  id: 'gpa-and-grade-tracker',
  name: 'GPA & Grade Tracker',
  isAI: true,
  sheetUrl: 'https://YOUR_GOOGLE_SHEET_LINK_HERE',
  etsy_title_keywords: ['gpa and grade tracker', 'gpa grade tracker', 'gpa-and-grade-tracker'],
  },
  'workout-tracker': {
  id: 'workout-tracker',
  name: 'Workout & Fitness Tracker',
  isAI: true,
  sheetUrl: 'https://docs.google.com/spreadsheets/d/1IgwDj3jqerRvg7riC9sBePopxRt96Y8g5_jyTcFS-1A/edit?usp=sharing',
  etsy_title_keywords: ['workout tracker', 'fitness tracker', 'workout-tracker', 'fitness-tracker'],
  },
  'four-year-planner': {
  id: 'four-year-planner',
  name: '4-Year College Planner',
  isAI: true,
  sheetUrl: 'https://docs.google.com/spreadsheets/d/1wf8RpVvRjr3fcnSuuAE1DPUqcxugvR-S-OvOsy5IVW4/edit?usp=sharing',
  etsy_title_keywords: ['four year planner', '4 year planner', 'college planner', 'four-year-planner', '4-year-planner'],
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
