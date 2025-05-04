/**
 * Constants
 */
const TITLE = 'Moceanic Calendar Checker';
const SUBTITLE = 'A tool to manage client sessions';
const PREFIX = 'C+ ';
const SEPARATOR = ' - ';
const RANGE_DAYS = {
	'1Y': { id: '1Y', name: 'A year', days: 365 },
	'6M': { id: '6M', name: 'Six months', days: 182 },
	'3M': { id: '3M', name: 'Three months', days: 91 },
	'1M': { id: '1M', name: 'One month', days: 30 }
};

/**
 * UI: Called whenever the user opens the add-on’s sidebar or clicks our icon
 */
function onHomepage(_) {
	return buildHomepageCard('3M');
}

/**
 * UI: home view
 */
function buildHomepageCard(rangeKey, debugMsg) {
	const now = new Date();
	const days = RANGE_DAYS[rangeKey].days || RANGE_DAYS['3M'].days;
	const msPerDay = 24 * 60 * 60 * 1000;
	const timeMin = new Date(now.getTime() - days * msPerDay).toISOString();
	const timeMax = new Date(now.getTime() + days * msPerDay).toISOString();

	const clients = fetchAllClients(timeMin, timeMax);
	const card = CardService.newCardBuilder().setHeader(
		CardService.newCardHeader().setTitle(TITLE).setSubtitle(SUBTITLE)
	);

	const filterSection = CardService.newCardSection();
	const dropdown = CardService.newSelectionInput()
		.setType(CardService.SelectionInputType.DROPDOWN)
		.setFieldName('timeRange')
		.setTitle('🔍 Filter by date range')
		.setOnChangeAction(CardService.newAction().setFunctionName('onTimeRangeChange'));

	Object.entries(RANGE_DAYS).forEach(([key, info]) => {
		dropdown.addItem(info.name, key, key === rangeKey);
	});

	filterSection.addWidget(dropdown);

	// if (debugMsg) {
	// 	section.addWidget(CardService.newTextParagraph().setText(`<b>DEBUG:</b> ${debugMsg}`));
	// }

	const clientsSection = CardService.newCardSection()
		.setHeader('👥 Clients')
		.addWidget(CardService.newTextParagraph().setText('Click a client to see their sessions<br><br>'));

	clients.forEach(({ name, topic, attendees, organization, events }) => {
		clientsSection.addWidget(
			CardService.newTextButton()
				.setText(`(${events}) ${organization}`)
				.setOnClickAction(
					CardService.newAction()
						.setFunctionName('showClientDetails')
						.setParameters({ name, topic, organization, attendees: JSON.stringify(attendees), timeMin, timeMax })
				)
		);
	});

	return card.addSection(filterSection).addSection(clientsSection).build();
}

/**
 * UI: client detail page
 */
function showClientDetails(event, timeMin, timeMax) {
	const { name, topic, attendees, organization } = event.parameters;
	const attendeesList = attendees ? JSON.parse(attendees) : [];
	const now = new Date();
	const past = [];
	const future = [];
	const events = fetchEventsForOrganization(organization, timeMin, timeMax);

	events.forEach(evt => {
		const start = new Date(evt.start.dateTime || evt.start.date);
		if (start < now) past.push(evt);
		else future.push(evt);
	});

	const card = CardService.newCardBuilder()
		.setHeader(CardService.newCardHeader().setTitle(`Client: ${organization}`))
		.addSection(
			CardService.newCardSection()
				.setHeader('Details')
				.addWidget(
					CardService.newTextParagraph().setText(
						`Topic: ${topic}<br>Name: ${name}<br>Attendees: ${attendeesList.length}<br>Events: ${past.length + future.length}`
					)
				)
		)
		.addSection(
			CardService.newCardSection()
				.setHeader(`<b>ATTENDEES</b> (${attendeesList.length})`)
				.setCollapsible(true)
				.addWidget(
					CardService.newTextParagraph().setText(
						`${attendeesList.map(({ name, email }) => (name != email ? `${name} (${email})` : `${email}`)).join('<br>')}`
					)
				)
		);

	[
		{
			section: CardService.newCardSection().setHeader(`<b>PAST EVENTS</b> (${past.length})`),
			list: past
		},
		{
			section: CardService.newCardSection().setHeader(`<b>FUTURE EVENTS</b> (${future.length})`),
			list: future
		}
	].forEach(group => {
		group.section.setCollapsible(group.list.length > 0);
		group.list.forEach(evt => {
			group.section.addWidget(
				CardService.newTextButton()
					.setText(evt.summary)
					.setOpenLink(CardService.newOpenLink().setUrl(evt.htmlLink).setOpenAs(CardService.OpenAs.FULL_SIZE))
			);
		});

		if (group.list.length == 0) {
			group.section.addWidget(CardService.newDecoratedText().setText(`- No events found -`));
		}
		card.addSection(group.section);
	});

	return card.build();
}

/**
 * Event handler: called whenever the user picks a different dropdown item
 */
function onTimeRangeChange(e) {
	const form = e.formInputs || {};
	const rangeKey = RANGE_DAYS[form.timeRange[0]].id || '3M';

	const newCard = buildHomepageCard(rangeKey);

	return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().updateCard(newCard)).build();
}

/**
 * Helper: fetch all distinct clients from our workshop events
 */
function fetchAllClients(timeMin, timeMax) {
	const cal = Calendar.Events;

	const events =
		cal.list('primary', {
			singleEvents: true,
			timeMin,
			timeMax,
			orderBy: 'startTime',
			q: PREFIX
		}).items || [];

	const clients = {};
	events.forEach(event => {
		let { topic, organization, name } = parseEventSummary(event.summary);
		clients[name] = clients[name] || { name, topic, organization, events: 0, attendees: [] };
		(event.attendees || []).forEach(attendee => {
			if (!attendee.self && attendee.responseStatus !== 'resource' && (event.summary || '').startsWith(PREFIX)) {
				clients[name].attendees.push({
					email: attendee.email,
					name: attendee.displayName || attendee.email
				});
			}
		});
		clients[name].events++;
	});

	return Object.values(clients).sort((a, b) => a.name.localeCompare(b.name));
}

/**
 * Helper: parse the event summary from oncehub
 */
function parseEventSummary(summary) {
	if (!summary.startsWith(PREFIX)) {
		console.error(`Invalid format: missing "${PREFIX}" prefix`);
		return { topic: '-', organization: '-', name: '-' };
	}

	const content = summary.slice(PREFIX.length);
	const firstSep = content.indexOf(SEPARATOR);
	const lastSep = content.lastIndexOf(SEPARATOR);

	if (firstSep < 0 || lastSep < 0 || firstSep === lastSep) {
		console.error(`Invalid format: expected two "${SEPARATOR}" separators`);
		return { topic: '-', organization: '-', name: '-' };
	}

	const topic = content.slice(0, firstSep).trim();
	const organization = content.slice(firstSep + SEPARATOR.length, lastSep).trim();
	const name = content.slice(lastSep + SEPARATOR.length).trim();

	return { topic, organization, name };
}

/**
 * Helper: fetch all events for a given client email
 */
function fetchEventsForClient(email, timeMin, timeMax) {
	const cal = Calendar.Events;

	return (
		cal.list('primary', {
			singleEvents: true,
			timeMin,
			timeMax,
			orderBy: 'startTime',
			q: email
		}).items || []
	).filter(evt => (evt.attendees || []).some(a => a.email === email));
}

/**
 * Helper: fetch all events for a given organization
 */
function fetchEventsForOrganization(organization, timeMin, timeMax) {
	const cal = Calendar.Events;

	const items =
		cal.list('primary', {
			singleEvents: true,
			timeMin,
			timeMax,
			orderBy: 'startTime',
			q: organization
		}).items || [];

	return items.filter(evt => {
		if (!evt.summary) return false;

		try {
			const { organization: orgInEvent } = parseEventSummary(evt.summary);
			return orgInEvent === organization;
		} catch (_) {
			return false;
		}
	});
}
