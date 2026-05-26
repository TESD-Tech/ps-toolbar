/**
 * Shared icon SVG map for PS Toolbar.
 *
 * Fill-style entries: exact SVGs from /ref/ (native viewBox, original colors).
 * Stroke-style entries: 24×24, stroke-width 1.5, round caps/joins.
 *
 * Add new icons here and to `availableIcons` in IconForm.svelte.
 */
export const iconSvgMap: Record<string, string> = {

  // ── Ref-folder originals (fill style, native viewBox) ──

  /** Communications — incoming envelope with arrow */
  mail: '<svg viewBox="0 0 32 32" fill="currentColor"><path d="M10.27 14a2.5 2.5 0 0 0-2.465 2.089l-2 12A2.5 2.5 0 0 0 8.27 31h17.958a2.5 2.5 0 0 0 2.466-2.089l2-12A2.5 2.5 0 0 0 28.23 14H10.271Zm-.493 2.418a.5.5 0 0 1 .494-.418h17.958a.5.5 0 0 1 .494.582l-.094.559l-10.29 5.54a.5.5 0 0 1-.523-.03l-8.129-5.69l.09-.543Zm-.454 2.729l5.16 3.611l-6.33 3.409l1.17-7.02Zm6.994 4.895l.352.247a2.5 2.5 0 0 0 2.62.153l.371-.2l7.023 4.469a.5.5 0 0 1-.454.289H8.271a.5.5 0 0 1-.482-.366l8.528-4.592Zm5.361-.887l6.535-3.519l-1.156 6.942l-5.379-3.423ZM1.75 16a.75.75 0 0 0 0 1.5h4.5a.75.75 0 0 0 0-1.5h-4.5Zm0 4a.75.75 0 0 0 0 1.5h3.5a.75.75 0 0 0 0-1.5h-3.5ZM1 24.75a.75.75 0 0 1 .75-.75h2.5a.75.75 0 0 1 0 1.5h-2.5a.75.75 0 0 1-.75-.75Z"/></svg>',

  /** Forms / documents — exact tesd_forms_header.svg (T/E Forms easter egg) */
  clipboard: '<svg version="1.1" viewBox="0 0 36 43.2" xmlns="http://www.w3.org/2000/svg"><style>.cf0{fill:#C1694F;stroke:#000;stroke-width:.5;stroke-miterlimit:10}.cf1{fill:#FFF;stroke:#000;stroke-width:.25;stroke-miterlimit:10}.cf2{fill:#CCD6DD;stroke:#000;stroke-width:.25;stroke-miterlimit:10}.cf3{fill:#292F33}.cf4{fill:#99AAB5}.cf5{font-family:ArialMT}.cf6{font-size:2.739px}</style><path class="cf0" d="M32,41c0,1.1-0.9,2-2,2H6c-1.1,0-2-0.9-2-2V14c0-1.1,0.9-2,2-2h24c1.1,0,2,0.9,2,2V41z"/><path class="cf1" d="M29,39c0,0.6-0.4,1-1,1H8c-0.6,0-1-0.4-1-1V16c0-0.6,0.4-1,1-1h20c0.6,0,1,0.4,1,1V39z"/><path class="cf2" d="M25,10h-4c0-1.7-1.3-3-3-3s-3,1.3-3,3h-4c-1.1,0-2,0.9-2,2v5h18v-5C27,10.9,26.1,10,25,10z"/><circle class="cf3" cx="18" cy="10" r="2"/><path class="cf4" d="M27,25c0,0.6-0.4,1-1,1H10c-0.6,0-1-0.4-1-1s0.4-1,1-1h16C26.6,24,27,24.4,27,25z M27,29c0,0.6-0.4,1-1,1H10c-0.6,0-1-0.4-1-1s0.4-1,1-1h16C26.6,28,27,28.4,27,29z M27,33c0,0.6-0.4,1-1,1H10c-0.6,0-1-0.4-1-1c0-0.6,0.4-1,1-1h16C26.6,32,27,32.4,27,33z M27,37c0,0.6-0.4,1-1,1h-9c-0.6,0-1-0.4-1-1c0-0.6,0.4-1,1-1h9C26.6,36,27,36.4,27,37z"/><text x="11.9" y="21.6" font-family="ArialMT" font-size="2.9px" fill="#292F33">T/E Forms</text></svg>',

  /** Athletics / activities */
  sports: '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M18.035 15.096a6.482 6.482 0 0 1-2.154 1.128c.078.496.119 1.005.119 1.524a8.003 8.003 0 0 0-.12-15.526A8.003 8.003 0 0 0 6.252 8c.518 0 1.027.04 1.523.119a6.482 6.482 0 0 1 1.128-2.154l9.131 9.131Zm1.06-1.06L15.062 10l.761-.761a4.23 4.23 0 0 0 2.428.76a4.23 4.23 0 0 0 2.22-.624a6.472 6.472 0 0 1-1.374 4.66ZM16.225 3.89c.66.24 1.27.585 1.811 1.014l-2.187 2.187A2.738 2.738 0 0 1 15.5 5.75c0-.717.274-1.37.724-1.86ZM14.76 8.178l-.76.761l-4.035-4.035a6.472 6.472 0 0 1 4.66-1.374A4.23 4.23 0 0 0 14 5.75c0 .903.281 1.74.761 2.428Zm5.349-.402a2.74 2.74 0 0 1-1.86.724c-.487 0-.944-.127-1.34-.349l2.186-2.186a6.49 6.49 0 0 1 1.014 1.81ZM4.25 10.5a.75.75 0 0 0-.75.75v2a7.25 7.25 0 0 0 7.25 7.25h2a.75.75 0 0 0 .75-.75v-2a7.25 7.25 0 0 0-7.25-7.25h-2ZM2 11.25A2.25 2.25 0 0 1 4.25 9h2A8.75 8.75 0 0 1 15 17.75v2A2.25 2.25 0 0 1 12.75 22h-2A8.75 8.75 0 0 1 2 13.25v-2Zm5.78 2.47a.75.75 0 0 0-1.06 1.06l2.5 2.5a.75.75 0 1 0 1.06-1.06l-2.5-2.5Z"/></svg>',

  /** Language / translation / ELL */
  language: '<svg viewBox="0 0 512 512" fill="currentColor"><path d="m478.33 433.6l-90-218a22 22 0 0 0-40.67 0l-90 218a22 22 0 1 0 40.67 16.79L316.66 406h102.67l18.33 44.39A22 22 0 0 0 458 464a22 22 0 0 0 20.32-30.4ZM334.83 362L368 281.65L401.17 362Zm-66.99-19.08a22 22 0 0 0-4.89-30.7c-.2-.15-15-11.13-36.49-34.73c39.65-53.68 62.11-114.75 71.27-143.49H330a22 22 0 0 0 0-44H214V70a22 22 0 0 0-44 0v20H54a22 22 0 0 0 0 44h197.25c-9.52 26.95-27.05 69.5-53.79 108.36c-31.41-41.68-43.08-68.65-43.17-68.87a22 22 0 0 0-40.58 17c.58 1.38 14.55 34.23 52.86 83.93c.92 1.19 1.83 2.35 2.74 3.51c-39.24 44.35-77.74 71.86-93.85 80.74a22 22 0 1 0 21.07 38.63c2.16-1.18 48.6-26.89 101.63-85.59c22.52 24.08 38 35.44 38.93 36.1a22 22 0 0 0 30.75-4.9Z"/></svg>',

  /** Health / emergency — 8-pointed star */
  health: '<svg viewBox="0 0 512 512" fill="currentColor"><path d="M351.9,255l108.1,-62.4l-48,-83.2l-108,62.4l0,-124.8l-96,0l0,124.8l-108,-62.4l-48,83.2l108.1,62.4l-108.1,62.4l48,83.2l108,-62.4l0,124.8l96,0l0,-124.8l108,62.4l48,-83.2z"/></svg>',

  /** Newly enrolled — user silhouette with plus */
  'people-plus': '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M15 14c-2.67 0-8 1.33-8 4v2h16v-2c0-2.67-5.33-4-8-4m-9-4V7H4v3H1v2h3v3h2v-3h3v-2m6 2a4 4 0 0 0 4-4a4 4 0 0 0-4-4a4 4 0 0 0-4 4a4 4 0 0 0 4 4Z"/></svg>',

  // ── Legacy stroke-style icons (24x24) ──

  /** Notifications */
  bell: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg>',
  /** Warnings */
  alert: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><path d="M12 9v4"/><circle cx="12" cy="17" r="1"/></svg>',
  /** Approvals / ratings */
  star: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></svg>',
  /** External links */
  external: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>',
  /** Single person / user */
  user: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>',
  /** Computer / tech labs */
  'computer-alt': '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="3" width="20" height="14" rx="2" ry="2"/><path d="M8 21h8"/><path d="M12 17v4"/></svg>',

  // ── 50 K-12 Education Icons (stroke style, 24x24) ──

  /** Subjects */
  math: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="3" width="16" height="18" rx="2"/><path d="M9 8h6"/><path d="M9 11l6 6"/><path d="M15 11l-6 6"/><path d="M9 19h6"/></svg>',
  science: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M9 3h6v5l5 11a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2l5-11V3z"/><path d="M9 3h6"/><path d="M7 13h10"/><path d="M7 17h10"/></svg>',
  book: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/><path d="M10 9h5"/><path d="M10 12h7"/><path d="M10 15h4"/></svg>',
  pencil: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M17 3a2.85 2.83 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z"/><path d="M15 5l4 4"/></svg>',
  paint: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="3"/><path d="M12 2v7"/><path d="M4.93 4.93l4.24 4.24"/><path d="M2 12h7"/><path d="M4.93 19.07l4.24-4.24"/><path d="M14.83 9.17L19.07 4.93"/></svg>',
  music: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M9 18V5l12-2v13"/><circle cx="6" cy="18" r="3"/><circle cx="18" cy="16" r="3"/></svg>',
  gym: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M6.5 6.5L17.5 17.5"/><path d="M8 3l13 13"/><path d="M3 8l13 13"/><path d="M3 3l2 2"/><path d="M19 19l2 2"/><path d="M3 21l2-2"/><path d="M19 3l2 2"/></svg>',
  globe: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="M2 12h20"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/></svg>',

  /** School logistics */
  bus: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="1" y="6" width="22" height="12" rx="3"/><path d="M5 18v2"/><path d="M19 18v2"/><circle cx="6" cy="16" r="1.5"/><circle cx="18" cy="16" r="1.5"/><path d="M15 6V4a1 1 0 0 0-1-1h-4a1 1 0 0 0-1 1v2"/><path d="M1 9h22"/></svg>',
  lunch: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M11 12H2v3a5 5 0 0 0 5 5h1"/><path d="M13 20h1a5 5 0 0 0 5-5v-3h-5"/><path d="M8 7v5"/><path d="M12 7v5"/><path d="M8 3v4"/><path d="M12 3v4"/></svg>',
  library: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z"/><path d="M12 6v7"/><path d="M11 9l2-2"/><path d="M13 9l-2-2"/></svg>',
  calendar: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4"/><path d="M8 2v4"/><path d="M3 10h18"/><path d="M8 14h.01"/><path d="M12 14h.01"/><path d="M16 14h.01"/><path d="M8 18h.01"/><path d="M12 18h.01"/><path d="M16 18h.01"/></svg>',
  clock: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>',
  grade: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01L12 2z"/><path d="M9 12l2 2 4-4"/></svg>',
  checklist: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg>',
  schedule: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><path d="M8 13h2"/><path d="M8 17h2"/><path d="M14 13h2"/><path d="M14 17h2"/></svg>',
  graduation: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M22 10v6M2 10l10-5 10 5-10 5z"/><path d="M6 12v5c3 3 9 3 12 0v-5"/></svg>',

  /** Support services */
  heart: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M20.84 4.61a5.5 5.5 0 0 0-7.78 0L12 5.67l-1.06-1.06a5.5 5.5 0 0 0-7.78 7.78l1.06 1.06L12 21.23l7.78-7.78 1.06-1.06a5.5 5.5 0 0 0 0-7.78z"/></svg>',
  nurse: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M9 12h6"/><path d="M12 9v6"/><path d="M12 22a10 10 0 1 0 0-20 10 10 0 0 0 0 20z"/></svg>',

  /** Communication */
  megaphone: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/><path d="M18 8a1 1 0 0 0-1-1H5"/></svg>',
  newspaper: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M4 22h16a2 2 0 0 0 2-2V4a2 2 0 0 0-2-2H8a2 2 0 0 0-2 2v16a2 2 0 0 1-4 0v-9"/><path d="M10 7h6"/><path d="M10 11h6"/><path d="M10 15h4"/></svg>',
  phone: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"/></svg>',

  /** People */
  family: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>',

  /** Facilities */
  building: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M3 21h18"/><path d="M9 21V3h6v18"/><path d="M3 7V3h4"/><path d="M21 7V3h-4"/><path d="M7 21h4"/><path d="M13 21h4"/></svg>',
  door: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M4 21h16"/><path d="M6 21V3h12v18"/><path d="M14 12h.01"/></svg>',
  playground: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M3 21h18"/><path d="M12 3v18"/><path d="M12 9l-4 5h8z"/><line x1="6" y1="14" x2="6" y2="21"/><line x1="18" y1="14" x2="18" y2="21"/></svg>',
  stage: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M3 21h18"/><path d="M3 10h18"/><path d="M5 6l7 4 7-4"/><path d="M5 10V6l7-4 7 4v4"/></svg>',

  /** Technology */
  tablet: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="4" y="2" width="16" height="20" rx="2"/><circle cx="12" cy="18" r="1"/></svg>',
  laptop: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="3" width="20" height="14" rx="2"/><path d="M2 17l1.5 2h17l1.5-2z"/><path d="M9 20h6"/></svg>',
  wifi: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.55a11 11 0 0 1 14.08 0"/><path d="M1.42 9a16 16 0 0 1 21.16 0"/><path d="M8.53 16.11a6 6 0 0 1 6.95 0"/><circle cx="12" cy="20" r="1"/></svg>',
  database: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"/><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/></svg>',

  /** Ideas & programs */
  lightbulb: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M9 18h6"/><path d="M10 22h4"/><path d="M15.09 14c.18-.98.65-1.74 1.41-2.5A4.65 4.65 0 0 0 18 8 6 6 0 0 0 6 8c0 1 .23 2.23 1.5 3.5A4.61 4.61 0 0 1 8.91 14"/></svg>',
  accessibility: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="5" r="2"/><path d="M7 11l5-3 5 3"/><path d="M12 8v7"/><path d="M9 22l3-7 3 7"/><path d="M5 14l4 2"/><path d="M15 16l4-2"/></svg>',
  gear: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>',

  /** Finance */
  dollar: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M12 1v22"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/></svg>',
  receipt: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M4 2v20l2-1 2 1 2-1 2 1 2-1 2 1 2-1 2 1V2l-2 1-2-1-2 1-2-1-2 1-2-1-2 1-2-1z"/><path d="M8 7h8"/><path d="M8 11h6"/><path d="M8 15h4"/></svg>',

  /** Achievement */
  award: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="6"/><path d="M15.477 12.89L17 22l-5-3-5 3 1.523-9.11"/></svg>',

  /** Safety */
  shield: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="M9 12l2 2 4-4"/></svg>',

  /** Behavior */
  'thumbs-up': '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 9V5a3 3 0 0 0-3-3l-4 9v11h11.28a2 2 0 0 0 2-1.7l1.38-9a2 2 0 0 0-2-2.3H14zM7 22H4a2 2 0 0 1-2-2v-7a2 2 0 0 1 2-2h3"/></svg>',

  /** Volunteer / support */
  hand: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M18 11V6a2 2 0 0 0-4 0v1M14 10V4a2 2 0 0 0-4 0v6M10 10.5V6a2 2 0 0 0-4 0v8"/><path d="M18 8a2 2 0 0 1 4 0v6a8 8 0 0 1-8 8h-2c-2.21 0-4.21-.9-5.66-2.34L3.5 14.66a2 2 0 0 1 .4-3.15 2 2 0 0 1 2.33.47L8 14"/></svg>',

  /** Goals */
  target: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/></svg>',

  /** Milestones */
  flag: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M4 15s1-1 4-1 5 2 8 2 4-1 4-1V3s-1 1-4 1-5-2-8-2-4 1-4 1z"/><path d="M4 22v-7"/></svg>',

  /** Access */
  key: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="7" cy="16" r="4"/><path d="M10 13l9-9"/><path d="M15 8l2-2"/><path d="M18 5l1-1"/></svg>',
  lock: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>',

  /** Communication */
  chat: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/><path d="M8 9h8"/><path d="M8 13h6"/></svg>',

  /** Collaboration */
  team: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>',

  /** Data */
  chart: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M18 20V10"/><path d="M12 20V4"/><path d="M6 20v-6"/></svg>',

  /** Upload / share */
  upload: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><path d="M12 3v12"/></svg>',

  /** Download / save */
  download: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><path d="M12 15V3"/></svg>',

  /** Camera / events / photos */
  camera: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M23 19a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h4l2-3h6l2 3h4a2 2 0 0 1 2 2z"/><circle cx="12" cy="13" r="4"/></svg>',
};

/**
 * Resolve an icon name to full SVG markup.
 * Normalizes input: lowercase, strips .svg suffix and path prefixes.
 */
export function resolveIconSvgPaths(name: string): string | null {
  const key = name.toLowerCase().replace(/\.svg$/, '').replace(/^.*[/\\]/, '');
  return iconSvgMap[key] ?? null;
}

/**
 * Alias — maintained for Toolbar.svelte compatibility.
 */
export function getIconSvgPaths(name: string): string | null {
  return resolveIconSvgPaths(name);
}