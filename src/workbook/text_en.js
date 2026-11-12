// SPDX-License-Identifier: AGPL-3.0-or-later
// English texts for the Darkroom Log workbook generator.
'use strict';
module.exports = {
  toc: {
    howto: 'How to use this workbook', processes: 'Processes · enlargers & lenses', papers: 'Printing papers',
    projidx: 'Project index', sessidx: 'Session index', learn: 'My darkroom learnings & notes', sessions: 'Darkroom sessions',
    qr: 'Quick reference', qr1: 'f-stop timing · print size change', qr2: 'RA-4 color casts · contrast grades',
    qr3: 'Zone system · printing rules', notes: 'General knowledge notes',
  },
  labels: { processes: 'PROCESSES', enlargers: 'ENLARGERS & LENSES', papers: 'PRINTING PAPERS', session: 'SESSION LOG' },
  howto: {
    intro: 'A log for darkroom printing: minimalist during your sessions – one block per session, one row per print. Write down what helps you, skip the rest. Setups and materials that rarely change are filled in once; afterwards, point to them with their short codes in the session logs. Use the quick references to look things up. The examples show one way to fill it in – the author\'s own notation and some variations of it. Adapt it to your own system and note it down below.',
    setupIntro: 'Fill in the setup tables once. In your sessions, you then only use their short codes (ABBREV).',
    setupExpl: {
      processes: 'one row per chemistry setup: chemical, dilution, time and temperature per bath. Its short code (abbrev) goes into CHEM in the session header.',
      enlargers: 'one row per enlarger & lens combination. Its short code (abbrev) goes into ENLARGER/LENS in the session header.',
      papersA5: 'one row per paper type and size. Use its short code (abbrev) in the frame header line of the session log.',
      papersA4: 'one row per paper type and size. Use its short code (abbrev) in the PAPER column of the session log.',
    },
    workflowTitle: 'A TYPICAL SESSION',
    workflow: [
      '▸ Before: check that process, enlarger and paper have a code in the setup tables.',
      '▸ During: write one row right after each print – only what changed.',
      '▸ After: add the session to the session index and mark your best prints.',
    ],
    sessionHead: 'Session header: date, project or title, enlarger/lens code and process code (CHEM). The number is pre-printed and links the block to the session index.',
    readTitle: 'READING THE EXAMPLE',
    read: [
      '▸ 1a: roll 12, frame 14, color head with M 40 / Y 20, 12 s at f/8 → too much contrast',
      '▸ 1b: same frame, softer filtration, 14 s; sky burned +4 s, bush down-left dodged 2 s → better, circled as best print',
      '▸ 2a, 2b: next frame, filter set at grade 2½, high intensity; 2b dodges the face 3 s and burns the corners +5 s → best print',
    ],
    notationTitle: 'NOTATION',
    sessionIntro: 'Session log: start each new frame with a frame header line across the table, then write one row per print. Leave a cell empty or mark it with " if nothing changed. Long sessions simply continue in the next block.',
    bulletsShort: [
      '▸ Frame header line:  — negative · film · paper · size',
      '▸ Negative ID: roll-frame, e.g. 12-14 (roll 12, frame 14)',
      '▸ Print ID: e.g. number = frame, letter = print (1a, 1b, 2a …)',
      '▸ Print back (pencil, on the dry print): date + print ID, e.g. 12.11.26 · 1b',
      '▸ FILTER: color head → C · M · Y; filter set / VC head → grade, across the columns',
    ],
    bulletX: '▸ Empty column after FILTER: label it as needed (light intensity, enlarger height …)',
    bulletINT: '▸ INT: light intensity setting of the enlarger head',
    bulletA4: '▸ A4: ROLL/NEG, PAPER and SIZE columns can replace the frame line',
    qr: 'Quick references at the back: f-stop printing times, how print size changes affect times, RA-4 color cast corrections, contrast grades and the zone system.',
    moreTitle: 'MORE EXAMPLES',
    morePrint: 'The extended how-to (2 pages, A4) explains the workbook in more detail, with more examples – scan the code to download it.',
    moreOnline: ['The ', { link: 'extended how-to' }, ' (2 pages, A4) explains the workbook in more detail, with more examples – click the link or scan the code.'],
    convTitle: 'MY CONVENTIONS',
    convText: 'Note your own ID schemes, abbreviations and habits here.',
  },
  // example entries (written in the handwriting font)
  samples: {
    // \n = line break (dilution on the first line, time/temperature on the second)
    processes: [['BW-1', 'Ilford Multigrade 1+9\n90s · 20°', 'Maco eco stop 1+19\n10s', 'Ilford Rapidfix 1+4\n2×45s', '5+ min', 'RC paper'],
      ['RA4', 'Bellini RA-4 kit, dev\n45s · 35°', 'stop\n30s', 'blix\n45s · 35°', '2 min', 'Jobo CPE-2, drum, 1 l']],
    enlargers: [['FM-1', 'Fujimoto G70', '35mm · color head', 'Rodenstock 50mm f/2.8', 'for 35mm']],
    enlargersExt: [['FM-1', 'Fujimoto G70', '35mm · color head', 'Rodenstock 50mm f/2.8', 'for 35mm'],
      ['DM-1', 'Durst M707BW', 'up to 6×7 · bw head', 'Minolta Rokkor 50mm f/4.5', 'use for 6×6, 6×7']],
    papers: [['FVM-1', 'Fomaspeed Variant 312 (matt)', '13×18', 'RC', 'B/W', 'box of 100, older']],
    papersExt: [['FVM-1', 'Fomaspeed Variant 312 (matt)', '13×18', 'RC', 'B/W', 'box of 100, older'],
      ['ARA-2', 'Adox RA-4 color paper, glossy', '20×25', 'RC', 'RA-4', 'keep cool']],
    // session example: header values, handwritten label of the free column (A5 only), rows (" = same as above)
    sessionHead: { no: '007', date: '12/11/26', title: 'travel Budapest - test prints', enl: 'FM-1', chem: 'BW-1' },
    xHead: 'INT',                             // handwritten label of the free column
    // {circ: '1b'} = print ID encircled (best print)
    sessionA5: [
      { line: '— 12-14 · Ilford HP5+ · FVM-1 · 13×18' },
      ['1a', '0', '40', '20', 'low', '12', '8', 'too much contrast'],
      [{ circ: '1b' }, '"', '25', '35', '"', '14', '"', 'sky +4s, bush down-left -2s'],
      { line: '— r12 / f17 · OptiMono 100 · FVM-1 · 10×15' },
      ['2a', { span: 3, t: 'grade 2½' }, 'high', '10', '11', 'filter'],
      [{ circ: '2b' }, { span: 3, t: '"' }, '"', '"', '"', 'dodge face 3s - burn corners +5s'],
    ],
    sessionA4: [
      ['12-14', '1a', '0', '40', '20', 'low', '12', '8', 'FVM-1', '13×18', 'too much contrast'],
      ['"', { circ: '1b' }, '"', '25', '35', '"', '14', '"', '"', '"', 'sky +4s, bush down-left -2s'],
      ['12-17', '2a', { span: 3, t: 'grade 2½' }, 'high', '10', '11', 'FVM-1', '10×15', 'filter set'],
      ['"', { circ: '2b' }, { span: 3, t: '"' }, '"', '"', '"', '"', '"', 'dodge face 3s - burn corners +5s'],
    ],
  },
  imprint: {
    lines: [
      'Darkroom Log – A Printing Workbook · Edition 1.0 · November 2026',
      '© 2026 Copyright by Hannes Walz (walz.analog@gmail.com)',
      '',
      'Licensed under CC BY-NC-SA 4.0 – creativecommons.org/licenses/by-nc-sa/4.0',
      'Printing and handing out copies in workshops and courses is explicitly allowed, including as part of a paid course, as long as the workbook is not sold separately from the course or workshop.',
    ],
    online: 'Source, updates and the extended how-to: github.com/wooolz/analog-workbook-darkroom-log',
    epubli: [
      'Herstellung: epubli – ein Service der neopubli GmbH, Köpenicker Straße 154a, 10997 Berlin',
      'Kontaktadresse nach EU-Produktsicherheitsverordnung: produktsicherheit@epubli.com',
    ],
  },
  ext: {
    repoUrl: 'https://github.com/wooolz/analog-workbook-darkroom-log',
    // stable download link of the latest release asset (upload the PDF with exactly this name to each release)
    howtoUrl: 'https://github.com/wooolz/analog-workbook-darkroom-log/releases/latest/download/darkroom-log_en_A4_howto-extended.pdf',
    intro: [
      'The Darkroom Log is a workbook for darkroom printing. It collects what you need to repeat a good print – or to understand why a print did not work: the session, the negative, the paper, filtration, time, aperture and every dodge and burn.',
      'There is no single right way to fill it in. The examples show the notation of the author or variants of it. Take what works for you, change the rest, and write your own conventions on the how-to page of your workbook.',
      ['If you feel comfortable using Python, you can change the code in the ', { link: 'github repository' }, ' and create your own, personalized version of this workbook.'],
    ],
    contents: [
      'Setup tables: note your processes, enlargers & lenses and printing papers – each with a short code.',
      'Session logs: minimalist note-taking during your sessions – one block per session, one row per print. Write down what helps you, skip the rest.',
      'Session index plus project index: find sessions again and know what belongs to which project. Note the best print, problems … The numbers match the log.',
      'Pages for your own learnings and notes: what you found out about papers, tools and techniques – so you do not have to find it out again.',
      'Quick references: f-stop timing, print size changes, RA-4 color casts, contrast grades (fill in the values for your head), zone system.',
    ],
    setup: 'Fill in the setup tables once, and check before each session that everything you need is in. Give every process, enlarger/lens combination and paper a short code (ABBREV) such as BW-1, RA4, FM-1 or P1. In the session log you then only write the code – the details stay on the setup pages.',
    session: [
      'Each session gets one block. Write the date, a title or project, the enlarger/lens code and the process code in the header. The pre-printed number identifies the session; long sessions simply continue in the next block (note "cont." in the title).',
      'Whenever you start a new frame, write a frame line across the table: negative, film, paper and print size. Below it, write one row per print. Leave cells empty or mark them with " when nothing changed – only the differences matter. In the A4 version, the ROLL/NEG, PAPER and SIZE columns can be used instead of the frame line.',
    ],
    ids: [
      'Negative ID: roll-frame, e.g. 12-14 = roll 12, frame 14. Number your films in the order you develop them.',
      'Print ID: a number for each frame in the session and a letter for each print of it: 1a, 1b, 1c, 2a …',
      'Back of the print: write date and print ID in pencil, e.g. 12.11.26 · 1b. With the session index, this leads straight back to the log entry.',
      'FILTER: with a color head, write the C · M · Y values. With a filter set or a variable contrast head, write the grade across the three columns.',
      'Empty column after FILTER: label it yourself – light intensity, enlarger height, timer program …',
    ],
    tips: [
      'Use a pencil – it writes on damp paper and is easy to read under the safelight.',
      'Write each entry right after the print, while you still remember what you did.',
      'Use short codes that speak to you: P1 for "paper 1" tells you less than FVM-1 for Foma Variant 312 (matt) in 13×18.',
      'Note failures: "too much contrast" saves time with the next print of the frame.',
      'Mark the best print of a frame, e.g. by circling the print number, so you spot it easily later.',
    ],
    footer: 'Darkroom Log · extended how-to · Edition 1.0 · November 2026 · CC BY-NC-SA 4.0',
  },
};
