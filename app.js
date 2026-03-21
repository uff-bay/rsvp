(() => {
  const config = window.RSVP_CONFIG || {};
  const form = document.getElementById('rsvpForm');
  const formMessage = document.getElementById('formMessage');
  const submitButton = document.getElementById('submitButton');
  const statusBanner = document.getElementById('statusBanner');
  const eventTitle = document.getElementById('eventTitle');
  const eventMeta = document.getElementById('eventMeta');
  const waiverA = document.getElementById('waiverA');
  const waiverB = document.getElementById('waiverB');
  const coordinatorEmail = document.getElementById('coordinatorEmail');
  const BLOCKED_EMAIL_DOMAINS = [
    'newarkunified.org',
    'fusdk12.net'
  ];

  if (!form) {
    return;
  }

  if (!config.webAppUrl || config.webAppUrl.includes('PASTE_YOUR_DEPLOYED_APPS_SCRIPT_EXEC_URL_HERE')) {
    formMessage.textContent = 'Setup is incomplete: add your Apps Script /exec URL in config.js.';
    submitButton.disabled = true;
    statusBanner.textContent = 'Setup is incomplete.';
    return;
  }

  form.action = config.webAppUrl;

  form.addEventListener('submit', (event) => {
    formMessage.textContent = '';

    const email = String(form.email.value || '').trim().toLowerCase();
    const emailConfirm = String(form.email_confirm.value || '').trim().toLowerCase();

    if (email !== emailConfirm) {
      event.preventDefault();
      formMessage.textContent = 'Your email address and confirmation email address must match.';
      form.email_confirm.focus();
      return;
    }

    if (isBlockedEmail(email)) {
      event.preventDefault();
      formMessage.textContent = 'School email addresses are not allowed. Please use a personal email.';
      form.email.focus();
      return;
    }

    submitButton.disabled = true;
    submitButton.textContent = 'Submitting…';
  });

  loadStatus();

  function loadStatus() {
    const callbackName = `__rsvpStatus_${Date.now()}_${Math.floor(Math.random() * 1000000)}`;
    const script = document.createElement('script');
    const separator = config.webAppUrl.includes('?') ? '&' : '?';

    window[callbackName] = (payload) => {
      cleanup();
      if (!payload || payload.ok !== true) {
        showLoadError();
        return;
      }
      applyStatus(payload);
    };

    script.src = `${config.webAppUrl}${separator}action=status&prefix=${encodeURIComponent(callbackName)}`;
    script.async = true;
    script.onerror = () => {
      cleanup();
      showLoadError();
    };
    document.body.appendChild(script);

    function cleanup() {
      delete window[callbackName];
      script.remove();
    }
  }

  function applyStatus(payload) {
    document.title = payload.event_title ? `${payload.event_title} RSVP` : 'RSVP';
    eventTitle.textContent = payload.event_title || 'RSVP';

    const metaParts = [payload.event_date_display, payload.event_location].filter(Boolean);
    eventMeta.textContent = metaParts.length ? metaParts.join(' • ') : 'Volunteer event';

    if (payload.waiver_url_a) {
      waiverA.href = payload.waiver_url_a;
    }
    if (payload.waiver_url_b) {
      waiverB.href = payload.waiver_url_b;
    }
    if (payload.coordinator_email) {
      coordinatorEmail.href = `mailto:${payload.coordinator_email}`;
      coordinatorEmail.textContent = payload.coordinator_email;
    }

    const openSpots = Number(payload.open_guaranteed_spots || 0);
    const confirmedCount = Number(payload.confirmed_count || 0);
    const waitlistCount = Number(payload.waitlist_count || 0);
    const capacity = Number(payload.guaranteed_capacity || 0);

    if (openSpots > 0) {
      statusBanner.className = 'status-banner ok';
      statusBanner.textContent = `${openSpots} guaranteed ${openSpots === 1 ? 'spot is' : 'spots are'} still available. ${confirmedCount} of ${capacity} guaranteed spots are filled.`;
      return;
    }

    statusBanner.className = 'status-banner waitlist';
    statusBanner.textContent = `All guaranteed spots are currently full. New RSVPs will join the waitlist. Current waitlist size: ${waitlistCount}.`;
  }

  function showLoadError() {
    statusBanner.className = 'status-banner';
    statusBanner.textContent = 'Could not load live availability right now. You can still submit the form.';
  }

  function isBlockedEmail(email) {
    const domain = email.split('@')[1]?.toLowerCase();
    return BLOCKED_EMAIL_DOMAINS.includes(domain);
  }
})();
