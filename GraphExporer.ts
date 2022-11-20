import { Client } from "@microsoft/microsoft-graph-client";
import { DateTime } from "luxon";
import { MyPluginSettings } from "main";

export default class GraphExplorer {
	graphClient: Client;
	settings: MyPluginSettings;

	constructor(graphClient: Client, settings: MyPluginSettings) {
		this.graphClient = graphClient;
		this.settings = settings;
	}

	public async obsidianRenderTodaysEvent() {
		const today = this.getToday();
		const tomorrow = this.getTomorrow();
		const events = await this.getUserEventsBetweenDates(today, tomorrow);

		const title = `## Events for the day of ${today.toLocaleDateString()} \n\n`;

		let eventsTable =
			title +
			`| Start | End | Subject | Notes | \n| --- | --- | --- | --- | \n`;

		events.value.forEach((event: any) => {
			eventsTable += `| ${this.getDateString(
				event.start.dateTime,
				event.start.timeZone,
				"time"
			)} | ${this.getDateString(
				event.end.dateTime,
				event.end.timeZone,
				"time"
			)} | ${
				event.subject
			} | ${this.createNoteLink(event)} \n`
		});

		return eventsTable;
	}

	public async obsidianRenderWeekEvents() {
		const { firstDayOfWeek, lastDayOfWeek } =
			this.getFirstAndLastDayOfWeek();

		const events = await this.getUserEventsBetweenDates(
			firstDayOfWeek,
			lastDayOfWeek
		);

		const title = `## Events for the week of ${firstDayOfWeek.toLocaleDateString()} to ${lastDayOfWeek.toLocaleDateString()} \n\n`;

		let eventsTable =
			title +
			`| Date | Start | End | Title | Notes | \n | --- | --- | --- | --- | --- |  \n`;

		events.value.forEach((event: any) => {
			eventsTable += `| ${this.getDateString(
				event.start.dateTime,
				event.start.timeZone,
				"date"
			)} | ${this.getDateString(
				event.start.dateTime,
				event.start.timeZone,
				"time"
			)} | ${this.getDateString(
				event.end.dateTime,
				event.end.timeZone,
				"time"
			)} | ${event.subject} | | \n`;
		});

		return eventsTable;
	}

	public async getUserEventsBetweenDates(firstDay: Date, lastDay: Date) {
		const events = await this.graphClient
			.api("/me/calendarview")
			.query({
				startDateTime: firstDay.toISOString(),
				endDateTime: lastDay.toISOString(),
			})
			.get();

		return events;
	}

	private getFirstAndLastDayOfWeek() {
		const today = new Date();
		const day = today.getDay();
		const diff = today.getDate() - day + (day == 0 ? -6 : 1); // adjust when day is sunday
		const firstDayOfWeek = new Date(today.setDate(diff));
		const lastDayOfWeek = new Date(today.setDate(diff + 6));

		this.resetDayToMidnight(firstDayOfWeek);
		this.resetDayToMidnightNextDay(lastDayOfWeek);

		return { firstDayOfWeek, lastDayOfWeek };
	}

	private resetDayToMidnight(date: Date) {
		date.setUTCHours(0, 0, 0, 0);
	}

	private resetDayToMidnightNextDay(date: Date) {
		date.setDate(date.getDate() + 1);
		date.setUTCHours(0, 0, 0, 0);
	}

	private getTomorrow() {
		const tomorrow = new Date();
		tomorrow.setDate(tomorrow.getDate() + 1);
		tomorrow.setUTCHours(0, 0, 0, 0);
		return tomorrow;
	}

	private getToday() {
		const today = new Date();
		today.setUTCHours(0, 0, 0, 0);
		return today;
	}

	private getDateString(
		date: string,
		timezone: string,
		format: "date" | "time"
	) {
		const d = DateTime.fromISO(date, { zone: timezone });

		const localTimeZone = DateTime.local().zoneName;

		const localTime = d.setZone(localTimeZone);

		switch (format) {
			case "date":
				return localTime.toFormat("yyMMdd");
			case "time":
				return localTime.toLocaleString(DateTime.TIME_24_SIMPLE);
			default:
				return localTime.toLocaleString(DateTime.DATE_SHORT);
		}
	}

	private getFolderName() {
		if (this.settings.folderName) return `${this.settings.folderName}`;
		else return "";
	}

	private createNoteTitle(event: any) {
		const date = this.getDateString(
			event.start.dateTime,
			event.start.timeZone,
			"date"
		);
		const subject = event.subject;
		return `${date} ${subject}.md`.trim().replace(/ /g, "_");
	}

	//[Test123](Meetings/test.md)

	private createNoteLink(event: any) {
		return `[${this.getDateString(
			event.start.dateTime,
			event.start.timeZone,
			"date"
		)} ${event.subject}](${this.getFolderName()}/${this.createNoteTitle(
			event
		)})`;
	}
}
