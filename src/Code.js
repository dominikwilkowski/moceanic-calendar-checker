/**
 * Called whenever the user opens the add-on’s sidebar or clicks our icon
 */
function onHomepage(e) {
	// Fetch your distinct client emails (or names) from all workshop events
	const clients = fetchAllClients();

	const card = CardService.newCardBuilder()
		.setHeader(
			CardService.newCardHeader().setTitle('Moceanic Calendar Checker').setSubtitle('Click a client to see sessions')
		)
		.build();

	const section = CardService.newCardSection();
	clients.forEach(({ email, name }) => {
		section.addWidget(
			CardService.newTextButton()
				.setText(name)
				.setOnClickAction(CardService.newAction().setFunctionName('showClientDetails').setParameters({ email: email }))
		);
	});

	return CardService.newCardBuilder().addSection(section).build();
}

/**
 * When a user clicks on a client button, show their past/future sessions
 */
function showClientDetails(e) {
	const email = e.parameters.email;
	const now = new Date();

	// Pull all events where this email is an attendee
	const events = fetchEventsForClient(email);
	const past = [],
		future = [];
	events.forEach(evt => {
		const start = new Date(evt.start.dateTime || evt.start.date);
		if (start < now) past.push(evt);
		else future.push(evt);
	});

	const section = CardService.newCardSection()
		.setHeader('Client: ' + email)
		.addWidget(CardService.newKeyValue().setTopLabel('Past sessions').setContent(past.length.toString()))
		.addWidget(CardService.newKeyValue().setTopLabel('Future sessions').setContent(future.length.toString()));

	// List each session with quick “Go to event” links
	[
		{ label: 'Past', list: past },
		{ label: 'Future', list: future }
	].forEach(group => {
		section.addWidget(CardService.newTextParagraph().setText(`<b>${group.label}:</b>`));
		group.list.forEach(evt => {
			section.addWidget(
				CardService.newTextButton()
					.setText(evt.summary + ' – ' + (evt.start.dateTime || evt.start.date))
					.setOpenLink(CardService.newOpenLink().setUrl(evt.htmlLink).setOpenAs(CardService.OpenAs.FULL_SIZE))
			);
		});
	});

	return CardService.newCardBuilder().addSection(section).build();
}

/**
 * Helper: fetch all distinct clients from our workshop events
 */
function fetchAllClients() {
	const cal = Calendar.Events;
	const now = new Date();
	const oneYearAgo = new Date(now.getTime() - 365 * 24 * 60 * 60 * 1000).toISOString();
	const oneYearAhead = new Date(now.getTime() + 365 * 24 * 60 * 60 * 1000).toISOString();

	const events =
		cal.list('primary', {
			singleEvents: true,
			timeMin: oneYearAgo,
			timeMax: oneYearAhead,
			orderBy: 'startTime'
		}).items || [];

	const clients = {};
	events.forEach(event => {
		(event.attendees || []).forEach(attendee => {
			if (!attendee.self && attendee.responseStatus !== 'resource' && (event.summary || '').startsWith('C+ ')) {
				clients[attendee.email] = { name: attendee.displayName || attendee.email, email: attendee.email };
			}
		});
	});

	return Object.values(clients).sort((a, b) => a.name.localeCompare(b.name));
}

/**
 * Helper: fetch all events for a given client email
 */
function fetchEventsForClient(email) {
	const cal = Calendar.Events;
	const now = new Date();
	const oneYearAgo = new Date(now.getTime() - 365 * 24 * 60 * 60 * 1000).toISOString();
	const oneYearAhead = new Date(now.getTime() + 365 * 24 * 60 * 60 * 1000).toISOString();

	return (
		cal.list('primary', {
			singleEvents: true,
			timeMin: oneYearAgo,
			timeMax: oneYearAhead,
			orderBy: 'startTime',
			q: email
		}).items || []
	).filter(evt => (evt.attendees || []).some(a => a.email === email));
}
