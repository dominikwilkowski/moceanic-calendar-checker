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
function onHomepage(e) {
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

	clients.forEach(({ name, topic, attendees, organization }) => {
		clientsSection.addWidget(
			CardService.newTextButton()
				.setText(organization)
				.setOnClickAction(
					CardService.newAction()
						.setFunctionName('showClientDetails')
						.setParameters({ name, topic, organization, timeMin, timeMax })
				)
		);
	});

	return card.addSection(filterSection).addSection(clientsSection).build();
}

/**
 * UI: client detail page
 */
function showClientDetails(event) {
	const { name, topic, organization } = event.parameters;
	const now = new Date();
	const past = [];
	const future = [];
	const events = fetchEventsForOrganization(organization);

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
				.addWidget(CardService.newTextParagraph().setText(`Topic: ${topic}<br>Name: ${name}`))
		);

	[
		{
			section: CardService.newCardSection().setHeader('<b>PAST SESSIONS</b>'),
			list: past
		},
		{
			section: CardService.newCardSection().setHeader('<b>FUTURE SESSIONS</b>'),
			list: future
		}
	].forEach(group => {
		group.section.addWidget(CardService.newKeyValue().setTopLabel('Sessions').setContent(group.list.length.toString()));
		let collapsibleSection = CardService.newCardSection()
			.setHeader('Events')
			.setCollapsible(group.list.length > 0);
		group.list.forEach(evt => {
			collapsibleSection.addWidget(
				CardService.newTextButton()
					.setText(evt.summary)
					.setOpenLink(CardService.newOpenLink().setUrl(evt.htmlLink).setOpenAs(CardService.OpenAs.FULL_SIZE))
			);
		});

		if (group.list.length == 0) {
			collapsibleSection.addWidget(CardService.newDecoratedText().setText(`- No events found -`));
		}
		card.addSection(group.section).addSection(collapsibleSection);
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
			orderBy: 'startTime'
		}).items || [];

	const clients = {};
	events.forEach(event => {
		(event.attendees || []).forEach(attendee => {
			if (!attendee.self && attendee.responseStatus !== 'resource' && (event.summary || '').startsWith('C+ ')) {
				let { topic, organization, name } = parseEventSummary(event.summary);

				if (clients[name]) {
					clients[name].attendees.push({ email: [attendee.email], name: attendee.displayName || attendee.email });
				} else {
					clients[name] = {
						name,
						topic,
						attendees: [{ email: [attendee.email], name: attendee.displayName || attendee.email }],
						organization
					};
				}
			}
		});
	});

	return Object.values(clients).sort((a, b) => a.name.localeCompare(b.name));
}

/**
 * Helper: parse the event summary from oncehub
 */
function parseEventSummary(summary) {
	if (!summary.startsWith(PREFIX)) {
		throw new Error(`Invalid format: missing "${PREFIX}" prefix`);
	}

	const content = summary.slice(PREFIX.length);
	const firstSep = content.indexOf(SEPARATOR);
	const lastSep = content.lastIndexOf(SEPARATOR);

	if (firstSep < 0 || lastSep < 0 || firstSep === lastSep) {
		throw new Error(`Invalid format: expected two "${SEPARATOR}" separators`);
	}

	const topic = content.slice(0, firstSep).trim();
	const organization = content.slice(firstSep + SEPARATOR.length, lastSep).trim();
	const name = content.slice(lastSep + SEPARATOR.length).trim();

	return { topic, organization, name };
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

/**
 * Helper: fetch all events for a given organization
 */
function fetchEventsForOrganization(organization) {
	const cal = Calendar.Events;
	const now = new Date();
	const oneYearAgo = new Date(now.getTime() - 365 * 24 * 60 * 60 * 1000).toISOString();
	const oneYearAhead = new Date(now.getTime() + 365 * 24 * 60 * 60 * 1000).toISOString();

	const items =
		cal.list('primary', {
			singleEvents: true,
			timeMin: oneYearAgo,
			timeMax: oneYearAhead,
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
