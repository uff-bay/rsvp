const SHEET_NAMES = {
  SETTINGS: 'Settings',
  DASHBOARD: 'Dashboard',
  RSVPS: 'RSVPs',
};

const SETTINGS_DEFAULTS = [
  ['event_title', 'Tree Planting Volunteer Event', 'Public event title'],
  ['event_date_display', 'Saturday, April 25, 2026 · 9:00 AM to 12:00 PM', 'Public date/time text'],
  ['event_location', 'Your park or neighborhood location', 'Public location text'],
  ['guaranteed_capacity', '50', 'Number of guaranteed spots before the waitlist starts'],
  ['coordinator_email', 'yooyoo@urbanforestfriends.org', 'Reply-to address shown to attendees'],
  ['public_form_url', 'https://YOUR-USERNAME.github.io/YOUR-REPO/', 'GitHub Pages URL for the RSVP page'],
  ['web_app_url', '', 'Paste the deployed Apps Script /exec URL here after deployment'],
  ['waiver_url_a', 'https://example.com/waiver-a.pdf', 'Public link to waiver A'],
  ['waiver_url_b', 'https://example.com/waiver-b.pdf', 'Public link to waiver B'],
  ['success_message', 'Thanks for RSVPing! Please watch your email for your confirmation and self-cancel link.', 'Shown after RSVP submission'],
  ['waitlist_message', 'You are on the waitlist. If a confirmed attendee withdraws or the organizer increases capacity, the next waitlisted person is promoted automatically.', 'Shown after waitlist signup'],
];

const DASHBOARD_KEYS = [
  'event_title',
  'guaranteed_capacity',
  'confirmed_count',
  'waitlist_count',
  'withdrawn_count',
  'open_guaranteed_spots',
  'last_refresh',
];

const RSVP_HEADERS = [
  'created_at',
  'updated_at',
  'rsvp_id',
  'cancel_token',
  'first_name',
  'last_name',
  'email',
  'participant_type',
  'organization_school',
  'tree_experience',
  'waiver_submission',
  'heard_about',
  'comments',
  'status',
  'waitlist_position',
  'status_reason',
  'confirmed_at',
  'promoted_at',
  'withdrawn_at',
  'last_email_status',
];

const STATUS = {
  CONFIRMED: 'confirmed',
  WAITLISTED: 'waitlisted',
  WITHDRAWN: 'withdrawn',
};

const PARTICIPANT_TYPES = new Set([
  'Student (age 14+)',
  'Adult',
]);

const TREE_EXPERIENCE_OPTIONS = new Set([
  'Yes, a lot',
  'Yes, a little',
  'No',
]);

const WAIVER_SUBMISSION_OPTIONS = new Set([
  'I will print them out physically and bring them to the event',
  'I will sign them digitally (or scan a copy) and send them to organizer',
  'I do not have the ability to print out a form or sign electronically, and request a set of forms',
]);

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Event RSVP')
    .addItem('Set up / repair workbook', 'setupWorkbook')
    .addItem('Increase guaranteed capacity...', 'increaseCapacityPrompt')
    .addItem('Refresh dashboard', 'refreshDashboard')
    .addToUi();
}

function setupWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  ensureSettingsSheet_(ss);
  ensureDashboardSheet_(ss);
  ensureRsvpSheet_(ss);
  refreshDashboard();

  SpreadsheetApp.getUi().alert(
    'Setup complete.',
    'The workbook is ready. Next: deploy this script as a web app, paste the /exec URL into Settings → web_app_url, then publish the GitHub Pages files.',
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

function refreshDashboard() {
  const ss = getSpreadsheet_();
  ensureDashboardSheet_(ss);
  const dashboard = ss.getSheetByName(SHEET_NAMES.DASHBOARD);
  const settings = readSettings_();
  const counts = computeCounts_();
  const values = [
    ['event_title', settings.event_title || ''],
    ['guaranteed_capacity', settings.guaranteed_capacity || '0'],
    ['confirmed_count', String(counts.confirmedCount)],
    ['waitlist_count', String(counts.waitlistCount)],
    ['withdrawn_count', String(counts.withdrawnCount)],
    ['open_guaranteed_spots', String(counts.openSpots)],
    ['last_refresh', nowIso_()],
  ];

  dashboard.getRange(2, 1, values.length, 2).setValues(values);
  autoSize_(dashboard, 2);
}

function increaseCapacityPrompt() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Increase guaranteed capacity',
    'How many additional guaranteed spots would you like to add?',
    ui.ButtonSet.OK_CANCEL,
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const increment = Number(String(response.getResponseText() || '').trim());
  if (!Number.isInteger(increment) || increment <= 0) {
    ui.alert('Please enter a whole number greater than 0.');
    return;
  }

  const result = increaseCapacityBy_(increment);
  ui.alert(
    'Capacity updated',
    `Guaranteed capacity is now ${result.newCapacity}. ${result.promotedCount} waitlisted ${result.promotedCount === 1 ? 'person was' : 'people were'} promoted automatically.`,
    ui.ButtonSet.OK,
  );
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  const action = String(params.action || '').trim();

  if (action === 'status') {
    return handleStatusRequest_(params);
  }

  if (action === 'cancel') {
    return renderCancelPage_(params);
  }

  return renderLandingPage_();
}

function doPost(e) {
  const params = (e && e.parameter) || {};
  const action = String(params.action || 'submit').trim();

  if (String(params.website || '').trim()) {
    return renderMessagePage_('Thanks!', '<p>Your request has been received.</p>');
  }

  if (action === 'submit') {
    return handleSubmitRequest_(params);
  }

  if (action === 'cancelConfirm') {
    return handleCancelConfirmRequest_(params);
  }

  return renderMessagePage_('Unknown action', '<p>That action is not supported.</p>');
}

function handleStatusRequest_(params) {
  const payload = publicStatusPayload_();
  const prefix = String(params.prefix || '').trim();

  if (prefix) {
    if (!/^[A-Za-z_$][0-9A-Za-z_$.]{0,100}$/.test(prefix)) {
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Invalid prefix' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(`${prefix}(${JSON.stringify(payload)})`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleSubmitRequest_(params) {
  let result;

  try {
    result = submitRsvp_(params);
  } catch (error) {
    return renderMessagePage_(
      'Could not process RSVP',
      `<p>${escapeHtml_(error.message || 'Something went wrong.')}</p>${backLinkHtml_()}`,
    );
  }

  const settings = readSettings_();
  const title = escapeHtml_(settings.event_title || 'Event RSVP');
  const backLink = backLinkHtml_();

  const emailHtml = result.emailStatus && String(result.emailStatus).startsWith('sent')
    ? `<p>An email was sent to <strong>${escapeHtml_(result.email)}</strong> with a link you can use to cancel your RSVP later.</p>`
    : `<p><strong>Heads up:</strong> we could not send the self-cancel email automatically (${escapeHtml_(String(result.emailStatus || 'unknown email issue'))}). Contact the organizer if you need help cancelling later.</p>`;

  if (result.status === STATUS.CONFIRMED) {
    return renderMessagePage_(
      'Your spot is guaranteed',
      `
      <p>Thanks, ${escapeHtml_(result.firstName)}. Your RSVP for <strong>${title}</strong> is confirmed.</p>
      <p>${escapeHtml_(settings.success_message || '')}</p>
      ${emailHtml}
      ${backLink}
      `,
    );
  }

  return renderMessagePage_(
    'You are on the waitlist',
    `
    <p>Thanks, ${escapeHtml_(result.firstName)}. All guaranteed spots are currently taken, so you have been added to the waitlist for <strong>${title}</strong>.</p>
    <p>Your current waitlist position is <strong>${result.waitlistPosition}</strong>.</p>
    <p>${escapeHtml_(settings.waitlist_message || '')}</p>
    ${emailHtml}
    ${backLink}
    `,
  );
}

function handleCancelConfirmRequest_(params) {
  const token = String(params.token || '').trim();
  if (!token) {
    return renderMessagePage_('Missing cancellation link', '<p>The cancellation link is invalid or incomplete.</p>');
  }

  const result = withdrawByToken_(token);
  const settings = readSettings_();
  const backLink = backLinkHtml_();

  if (result.outcome === 'already_withdrawn') {
    return renderMessagePage_(
      'Already cancelled',
      `<p>This RSVP was already withdrawn earlier.</p>${backLink}`,
    );
  }

  if (result.outcome === 'not_found') {
    return renderMessagePage_(
      'Cancellation link not found',
      `<p>This cancellation link is no longer valid. Contact the organizer if you need help.</p>${backLink}`,
    );
  }

  const promotionHtml = result.promotedPerson
    ? `<p><strong>${escapeHtml_(result.promotedPerson)}</strong> was automatically promoted from the waitlist.</p>`
    : '';

  return renderMessagePage_(
    'RSVP cancelled',
    `
    <p>Your RSVP for <strong>${escapeHtml_(settings.event_title || 'the event')}</strong> has been cancelled.</p>
    <p>Status before cancellation: <strong>${escapeHtml_(result.previousStatus)}</strong>.</p>
    ${promotionHtml}
    ${backLink}
    `,
  );
}

function submitRsvp_(params) {
  const data = normalizeSubmission_(params);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const settings = readSettings_();
    const capacity = positiveInt_(settings.guaranteed_capacity, 0);
    const sheet = getRsvpSheet_();
    const records = readRsvpRecords_();

    const duplicate = records.find((record) =>
      record.status !== STATUS.WITHDRAWN &&
      record.emailLower === data.email.toLowerCase() &&
      record.firstNameLower === data.firstName.toLowerCase() &&
      record.lastNameLower === data.lastName.toLowerCase()
    );

    if (duplicate) {
      throw new Error('It looks like this person is already RSVP’d with the same email address. Use the cancellation link from the email you already received, or contact the organizer.');
    }

    const counts = computeCountsFromRecords_(records, capacity);
    const now = nowIso_();
    const rsvpId = Utilities.getUuid();
    const cancelToken = createCancelToken_();
    const status = counts.confirmedCount < capacity ? STATUS.CONFIRMED : STATUS.WAITLISTED;
    const waitlistPosition = status === STATUS.WAITLISTED ? counts.waitlistCount + 1 : '';

    const row = [
      now,
      now,
      rsvpId,
      cancelToken,
      data.firstName,
      data.lastName,
      data.email,
      data.participantType,
      data.organizationSchool,
      data.treeExperience,
      data.waiverSubmission,
      data.heardAbout,
      data.comments,
      status,
      waitlistPosition,
      status === STATUS.CONFIRMED ? 'initial guaranteed spot' : 'initial waitlist placement',
      status === STATUS.CONFIRMED ? now : '',
      '',
      '',
      '',
    ];

    sheet.appendRow(row);
    const appendedRowNumber = sheet.getLastRow();

    const emailStatus = sendSubmissionEmail_({
      firstName: data.firstName,
      lastName: data.lastName,
      email: data.email,
      status,
      waitlistPosition,
      cancelToken,
    }, settings);

    setCellByHeader_(sheet, appendedRowNumber, 'last_email_status', emailStatus);
    refreshDashboard();

    return {
      status,
      waitlistPosition,
      firstName: data.firstName,
      email: data.email,
      emailStatus,
    };
  } finally {
    lock.releaseLock();
  }
}

function withdrawByToken_(token) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = getRsvpSheet_();
    const record = findRecordByToken_(token);
    if (!record) {
      return { outcome: 'not_found' };
    }

    if (record.status === STATUS.WITHDRAWN) {
      return { outcome: 'already_withdrawn' };
    }

    const now = nowIso_();
    setCellByHeader_(sheet, record.rowNumber, 'status', STATUS.WITHDRAWN);
    setCellByHeader_(sheet, record.rowNumber, 'status_reason', 'self-withdrawn');
    setCellByHeader_(sheet, record.rowNumber, 'withdrawn_at', now);
    setCellByHeader_(sheet, record.rowNumber, 'waitlist_position', '');
    setCellByHeader_(sheet, record.rowNumber, 'updated_at', now);

    renumberWaitlist_();
    const promotion = promoteWaitlistIntoOpenSpots_();
    refreshDashboard();

    return {
      outcome: 'withdrawn',
      previousStatus: record.status,
      promotedPerson: promotion.promotedNames[0] || '',
    };
  } finally {
    lock.releaseLock();
  }
}

function increaseCapacityBy_(increment) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const settings = readSettings_();
    const current = positiveInt_(settings.guaranteed_capacity, 0);
    const nextCapacity = current + increment;
    setSetting_('guaranteed_capacity', String(nextCapacity));

    const promotion = promoteWaitlistIntoOpenSpots_();
    refreshDashboard();

    return {
      newCapacity: nextCapacity,
      promotedCount: promotion.promotedCount,
    };
  } finally {
    lock.releaseLock();
  }
}

function promoteWaitlistIntoOpenSpots_() {
  const settings = readSettings_();
  const capacity = positiveInt_(settings.guaranteed_capacity, 0);
  const sheet = getRsvpSheet_();
  const promotedNames = [];
  let promotedCount = 0;

  while (true) {
    const records = readRsvpRecords_();
    const counts = computeCountsFromRecords_(records, capacity);
    if (counts.confirmedCount >= capacity) {
      break;
    }

    const next = records
      .filter((record) => record.status === STATUS.WAITLISTED)
      .sort(compareWaitlistRecords_)[0];

    if (!next) {
      break;
    }

    const now = nowIso_();
    setCellByHeader_(sheet, next.rowNumber, 'status', STATUS.CONFIRMED);
    setCellByHeader_(sheet, next.rowNumber, 'waitlist_position', '');
    setCellByHeader_(sheet, next.rowNumber, 'status_reason', 'promoted from waitlist');
    setCellByHeader_(sheet, next.rowNumber, 'confirmed_at', now);
    setCellByHeader_(sheet, next.rowNumber, 'promoted_at', now);
    setCellByHeader_(sheet, next.rowNumber, 'updated_at', now);

    const emailStatus = sendPromotionEmail_(next, settings);
    setCellByHeader_(sheet, next.rowNumber, 'last_email_status', emailStatus);

    promotedCount += 1;
    promotedNames.push(`${next.firstName} ${next.lastName}`.trim());
  }

  renumberWaitlist_();

  return {
    promotedCount,
    promotedNames,
  };
}

function renderCancelPage_(params) {
  const token = String(params.token || '').trim();
  if (!token) {
    return renderMessagePage_('Missing cancellation link', '<p>The cancellation link is invalid or incomplete.</p>');
  }

  const record = findRecordByToken_(token);
  const settings = readSettings_();

  if (!record) {
    return renderMessagePage_('Cancellation link not found', '<p>This cancellation link is no longer valid. Contact the organizer if you need help.</p>');
  }

  if (record.status === STATUS.WITHDRAWN) {
    return renderMessagePage_('Already cancelled', '<p>This RSVP was already cancelled earlier.</p>');
  }

  const actionUrl = escapeHtml_(settings.web_app_url || '');
  const content = `
    <p>You are about to cancel the RSVP below:</p>
    <div class="card">
      <p><strong>Name:</strong> ${escapeHtml_(`${record.firstName} ${record.lastName}`.trim())}</p>
      <p><strong>Email:</strong> ${escapeHtml_(record.email)}</p>
      <p><strong>Current status:</strong> ${escapeHtml_(record.status)}</p>
      ${record.waitlistPosition ? `<p><strong>Waitlist position:</strong> ${escapeHtml_(String(record.waitlistPosition))}</p>` : ''}
    </div>
    <form method="post" action="${actionUrl}">
      <input type="hidden" name="action" value="cancelConfirm">
      <input type="hidden" name="token" value="${escapeHtml_(token)}">
      <button class="danger" type="submit">Yes, cancel my RSVP</button>
    </form>
    ${backLinkHtml_()}
  `;

  return renderMessagePage_('Cancel RSVP', content);
}

function renderLandingPage_() {
  const settings = readSettings_();
  const publicFormUrl = String(settings.public_form_url || '').trim();
  const content = publicFormUrl
    ? `<p>This RSVP backend is live.</p><p><a class="button-link" href="${escapeHtml_(publicFormUrl)}">Open the public RSVP page</a></p>`
    : '<p>This RSVP backend is live. Set <code>public_form_url</code> in the Settings sheet to show a return link here.</p>';

  return renderMessagePage_('RSVP backend ready', content);
}

function publicStatusPayload_() {
  const settings = readSettings_();
  const counts = computeCounts_();

  return {
    ok: true,
    event_title: settings.event_title || '',
    event_date_display: settings.event_date_display || '',
    event_location: settings.event_location || '',
    coordinator_email: settings.coordinator_email || '',
    waiver_url_a: settings.waiver_url_a || '',
    waiver_url_b: settings.waiver_url_b || '',
    guaranteed_capacity: positiveInt_(settings.guaranteed_capacity, 0),
    confirmed_count: counts.confirmedCount,
    waitlist_count: counts.waitlistCount,
    withdrawn_count: counts.withdrawnCount,
    open_guaranteed_spots: counts.openSpots,
    waitlist_open: counts.openSpots <= 0,
    public_form_url: settings.public_form_url || '',
  };
}

function computeCounts_() {
  const settings = readSettings_();
  const capacity = positiveInt_(settings.guaranteed_capacity, 0);
  const records = readRsvpRecords_();
  return computeCountsFromRecords_(records, capacity);
}

function computeCountsFromRecords_(records, capacity) {
  const confirmedCount = records.filter((record) => record.status === STATUS.CONFIRMED).length;
  const waitlistCount = records.filter((record) => record.status === STATUS.WAITLISTED).length;
  const withdrawnCount = records.filter((record) => record.status === STATUS.WITHDRAWN).length;

  return {
    confirmedCount,
    waitlistCount,
    withdrawnCount,
    openSpots: Math.max(capacity - confirmedCount, 0),
  };
}

function normalizeSubmission_(params) {
  const firstName = cleanText_(params.first_name);
  const lastName = cleanText_(params.last_name);
  const email = cleanEmail_(params.email);
  const emailConfirm = cleanEmail_(params.email_confirm);
  const participantType = cleanText_(params.participant_type);
  const organizationSchool = cleanText_(params.organization_school);
  const treeExperience = cleanText_(params.tree_experience);
  const waiverSubmission = cleanText_(params.waiver_submission);
  const heardAbout = cleanText_(params.heard_about);
  const comments = cleanText_(params.comments);

  if (!firstName || !lastName) {
    throw new Error('Please provide both a first name and a last name.');
  }

  if (!email || !/^\S+@\S+\.\S+$/.test(email)) {
    throw new Error('Please provide a valid email address.');
  }

  if (email !== emailConfirm) {
    throw new Error('Your email address and confirmation email address do not match.');
  }

  if (!PARTICIPANT_TYPES.has(participantType)) {
    throw new Error('Please choose whether you are a student or an adult.');
  }

  if (!TREE_EXPERIENCE_OPTIONS.has(treeExperience)) {
    throw new Error('Please choose a tree-planting experience option.');
  }

  if (!WAIVER_SUBMISSION_OPTIONS.has(waiverSubmission)) {
    throw new Error('Please choose how you will submit the waiver forms.');
  }

  return {
    firstName,
    lastName,
    email,
    participantType,
    organizationSchool,
    treeExperience,
    waiverSubmission,
    heardAbout,
    comments,
  };
}

function sendSubmissionEmail_(record, settings) {
  if (!record.email) {
    return 'no_email';
  }

  if (!settings.web_app_url) {
    return 'missing_web_app_url';
  }

  if (MailApp.getRemainingDailyQuota() < 1) {
    return 'email_quota_exhausted';
  }

  const cancelUrl = buildCancelUrl_(settings.web_app_url, record.cancelToken);
  const eventTitle = settings.event_title || 'Event RSVP';
  const dateText = settings.event_date_display || '';
  const locationText = settings.event_location || '';
  const replyTo = settings.coordinator_email || undefined;
  const subject = record.status === STATUS.CONFIRMED
    ? `Confirmed RSVP: ${eventTitle}`
    : `Waitlist RSVP: ${eventTitle}`;

  const intro = record.status === STATUS.CONFIRMED
    ? 'Your spot is guaranteed.'
    : `You are currently on the waitlist${record.waitlistPosition ? ` at position ${record.waitlistPosition}` : ''}.`;

  const htmlBody = `
    <p>Hi ${escapeHtml_(record.firstName)},</p>
    <p>Thanks for RSVPing for <strong>${escapeHtml_(eventTitle)}</strong>.</p>
    <p><strong>${escapeHtml_(intro)}</strong></p>
    ${dateText ? `<p><strong>When:</strong> ${escapeHtml_(dateText)}</p>` : ''}
    ${locationText ? `<p><strong>Where:</strong> ${escapeHtml_(locationText)}</p>` : ''}
    <p>You can cancel your RSVP later with this link:</p>
    <p><a href="${escapeHtml_(cancelUrl)}">Manage or cancel my RSVP</a></p>
    <p>This is a two-click flow: open the link, then press the cancel button if you really want to withdraw.</p>
  `;

  const body = [
    `Hi ${record.firstName},`,
    '',
    `Thanks for RSVPing for ${eventTitle}.`,
    intro,
    dateText ? `When: ${dateText}` : '',
    locationText ? `Where: ${locationText}` : '',
    '',
    `Manage or cancel your RSVP: ${cancelUrl}`,
  ].filter(Boolean).join('\n');

  MailApp.sendEmail({
    to: record.email,
    subject,
    body,
    htmlBody,
    name: eventTitle,
    replyTo,
  });

  return `sent ${nowIso_()}`;
}

function sendPromotionEmail_(record, settings) {
  if (!record.email) {
    return 'no_email';
  }

  if (!settings.web_app_url) {
    return 'missing_web_app_url';
  }

  if (MailApp.getRemainingDailyQuota() < 1) {
    return 'email_quota_exhausted';
  }

  const cancelUrl = buildCancelUrl_(settings.web_app_url, record.cancelToken);
  const eventTitle = settings.event_title || 'Event RSVP';
  const dateText = settings.event_date_display || '';
  const locationText = settings.event_location || '';
  const replyTo = settings.coordinator_email || undefined;

  const htmlBody = `
    <p>Hi ${escapeHtml_(record.firstName)},</p>
    <p>Good news — a guaranteed spot opened up for <strong>${escapeHtml_(eventTitle)}</strong>, and you have been moved off the waitlist.</p>
    ${dateText ? `<p><strong>When:</strong> ${escapeHtml_(dateText)}</p>` : ''}
    ${locationText ? `<p><strong>Where:</strong> ${escapeHtml_(locationText)}</p>` : ''}
    <p>If you no longer want the spot, you can cancel here:</p>
    <p><a href="${escapeHtml_(cancelUrl)}">Manage or cancel my RSVP</a></p>
  `;

  const body = [
    `Hi ${record.firstName},`,
    '',
    `A guaranteed spot opened up for ${eventTitle}, and you have been moved off the waitlist.`,
    dateText ? `When: ${dateText}` : '',
    locationText ? `Where: ${locationText}` : '',
    '',
    `Manage or cancel your RSVP: ${cancelUrl}`,
  ].filter(Boolean).join('\n');

  MailApp.sendEmail({
    to: record.email,
    subject: `You are off the waitlist: ${eventTitle}`,
    body,
    htmlBody,
    name: eventTitle,
    replyTo,
  });

  return `promotion_email_sent ${nowIso_()}`;
}

function ensureSettingsSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
  }

  const needsHeader = sheet.getLastRow() === 0 || String(sheet.getRange(1, 1).getValue()) !== 'Key';
  if (needsHeader) {
    sheet.clear();
    sheet.getRange(1, 1, 1, 3).setValues([['Key', 'Value', 'Notes']]);
    sheet.setFrozenRows(1);
  }

  const existing = readSettingsMap_(sheet);
  const rowsToAppend = SETTINGS_DEFAULTS.filter(([key]) => !(key in existing));
  if (rowsToAppend.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 3).setValues(rowsToAppend);
  }

  styleHeader_(sheet, 3);
  autoSize_(sheet, 3);
}

function ensureDashboardSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.DASHBOARD);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.DASHBOARD);
  }

  sheet.clear();
  sheet.getRange(1, 1, 1, 2).setValues([['Metric', 'Value']]);
  const rows = DASHBOARD_KEYS.map((key) => [key, '']);
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  sheet.setFrozenRows(1);
  styleHeader_(sheet, 2);
  autoSize_(sheet, 2);
}

function ensureRsvpSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.RSVPS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.RSVPS);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, RSVP_HEADERS.length).setValues([RSVP_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, RSVP_HEADERS.length).createFilter();
  } else {
    const currentHeaders = sheet.getRange(1, 1, 1, RSVP_HEADERS.length).getValues()[0];
    const mismatch = RSVP_HEADERS.some((header, index) => currentHeaders[index] !== header);
    if (mismatch) {
      throw new Error('The RSVPs sheet headers do not match the expected structure. Make a backup of the sheet, then run setup again on a clean workbook.');
    }
  }

  styleHeader_(sheet, RSVP_HEADERS.length);
  autoSize_(sheet, RSVP_HEADERS.length);
}

function renumberWaitlist_() {
  const sheet = getRsvpSheet_();
  const records = readRsvpRecords_()
    .filter((record) => record.status === STATUS.WAITLISTED)
    .sort(compareWaitlistRecords_);

  records.forEach((record, index) => {
    setCellByHeader_(sheet, record.rowNumber, 'waitlist_position', index + 1);
    setCellByHeader_(sheet, record.rowNumber, 'updated_at', nowIso_());
  });
}

function compareWaitlistRecords_(a, b) {
  const posA = Number(a.waitlistPosition || Number.MAX_SAFE_INTEGER);
  const posB = Number(b.waitlistPosition || Number.MAX_SAFE_INTEGER);
  if (posA !== posB) {
    return posA - posB;
  }
  return String(a.createdAt).localeCompare(String(b.createdAt));
}

function readRsvpRecords_() {
  const sheet = getRsvpSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  const values = sheet.getRange(2, 1, lastRow - 1, RSVP_HEADERS.length).getValues();
  return values.map((row, index) => mapRowToRecord_(row, index + 2));
}

function mapRowToRecord_(row, rowNumber) {
  const record = { rowNumber };
  RSVP_HEADERS.forEach((header, index) => {
    record[camelCase_(header)] = row[index];
  });
  record.emailLower = String(record.email || '').trim().toLowerCase();
  record.firstNameLower = String(record.firstName || '').trim().toLowerCase();
  record.lastNameLower = String(record.lastName || '').trim().toLowerCase();
  return record;
}

function findRecordByToken_(token) {
  const normalized = String(token || '').trim();
  if (!normalized) {
    return null;
  }

  return readRsvpRecords_().find((record) => String(record.cancelToken || '') === normalized) || null;
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) {
    return SpreadsheetApp.openById(id);
  }

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', active.getId());
    return active;
  }

  throw new Error('Spreadsheet not configured yet. Open the bound sheet and run Event RSVP → Set up / repair workbook.');
}

function getRsvpSheet_() {
  const sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.RSVPS);
  if (!sheet) {
    throw new Error('The RSVPs sheet is missing. Run Event RSVP → Set up / repair workbook.');
  }
  return sheet;
}

function readSettings_() {
  const sheet = getSpreadsheet_().getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) {
    throw new Error('The Settings sheet is missing. Run Event RSVP → Set up / repair workbook.');
  }

  const map = readSettingsMap_(sheet);
  const obj = {};
  Object.keys(map).forEach((key) => {
    obj[key] = map[key].value;
  });
  return obj;
}

function readSettingsMap_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {};
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const out = {};
  rows.forEach((row, index) => {
    const key = String(row[0] || '').trim();
    if (!key) {
      return;
    }
    out[key] = {
      value: row[1],
      notes: row[2],
      rowNumber: index + 2,
    };
  });
  return out;
}

function setSetting_(key, value) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  const map = readSettingsMap_(sheet);
  if (map[key]) {
    sheet.getRange(map[key].rowNumber, 2).setValue(value);
    return;
  }
  sheet.appendRow([key, value, '']);
}

function setCellByHeader_(sheet, rowNumber, header, value) {
  const columnIndex = RSVP_HEADERS.indexOf(header);
  if (columnIndex === -1) {
    throw new Error(`Unknown header: ${header}`);
  }
  sheet.getRange(rowNumber, columnIndex + 1).setValue(value);
}

function buildCancelUrl_(webAppUrl, token) {
  const separator = webAppUrl.indexOf('?') === -1 ? '?' : '&';
  return `${webAppUrl}${separator}action=cancel&token=${encodeURIComponent(token)}`;
}

function createCancelToken_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function renderMessagePage_(title, bodyHtml) {
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>${escapeHtml_(title)}</title>
        <style>
          body { font-family: Arial, sans-serif; background: #f6f8fb; color: #12202f; margin: 0; padding: 24px; }
          .wrap { max-width: 760px; margin: 0 auto; }
          .panel { background: #fff; border: 1px solid #d7dfeb; border-radius: 16px; padding: 24px; box-shadow: 0 8px 30px rgba(18,32,47,0.07); }
          h1 { margin-top: 0; font-size: 1.9rem; }
          p { line-height: 1.6; }
          .card { background: #f8fbff; border: 1px solid #d7e4f0; border-radius: 12px; padding: 16px; margin: 16px 0; }
          button, .button-link { display: inline-block; padding: 12px 18px; border-radius: 10px; border: 0; font-size: 1rem; text-decoration: none; background: #1f6feb; color: #fff; cursor: pointer; }
          .danger { background: #c0362c; }
          .back-link { margin-top: 16px; }
          code { background: #eef3f9; padding: 2px 6px; border-radius: 6px; }
        </style>
      </head>
      <body>
        <div class="wrap">
          <div class="panel">
            <h1>${escapeHtml_(title)}</h1>
            ${bodyHtml}
          </div>
        </div>
      </body>
    </html>
  `;

  return HtmlService.createHtmlOutput(html).setTitle(title);
}

function backLinkHtml_() {
  const settings = readSettings_();
  const publicFormUrl = String(settings.public_form_url || '').trim();
  if (!publicFormUrl) {
    return '';
  }
  return `<p class="back-link"><a class="button-link" href="${escapeHtml_(publicFormUrl)}">Back to RSVP page</a></p>`;
}

function styleHeader_(sheet, columns) {
  sheet.getRange(1, 1, 1, columns)
    .setFontWeight('bold')
    .setBackground('#dce9f8');
}

function autoSize_(sheet, columns) {
  for (let i = 1; i <= columns; i += 1) {
    sheet.autoResizeColumn(i);
  }
}

function nowIso_() {
  return new Date().toISOString();
}

function positiveInt_(value, fallback) {
  const num = Number(value);
  return Number.isFinite(num) && num >= 0 ? Math.floor(num) : fallback;
}

function cleanText_(value) {
  return String(value || '').trim().replace(/\s+/g, ' ');
}

function cleanEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function camelCase_(text) {
  return String(text).replace(/_([a-z])/g, (_, chr) => chr.toUpperCase());
}
