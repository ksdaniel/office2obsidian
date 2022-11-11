import { Client } from "@microsoft/microsoft-graph-client";
import { DateTime } from "luxon";

export default class GraphExplorer {
	graphClient: Client;

	constructor(graphClient: Client) {
		this.graphClient = graphClient;
	}

	public async obsidianRenderDayGraphEventsAsTable() {
		const events = await this.getUserEventsForToday();
		const headline = `### Events for ${DateTime.local().toLocaleString(
			DateTime.DATE_FULL
		)}\n\n`;
		let output = headline + "| Subject | Start | End | Done |\n";
		output += "| --- | --- | --- | --- |\n";
		for (let event of events.value) {
			output += `| ${event.subject} | ${DateTime.fromISO(
				event.start.dateTime
			).toLocaleString(DateTime.TIME_24_SIMPLE)} | ${DateTime.fromISO(
				event.end.dateTime
			).toLocaleString(DateTime.TIME_24_SIMPLE)} |  |\n`;
		}
		return output;
	}

	public async obsidianRenderWeekGraphEventsAsTable() {
		const events = await this.getUserEventsForWeek();
		const headline = `### Events for ${DateTime.local().toLocaleString(
			DateTime.DATE_FULL
		)} - ${DateTime.local()
			.plus({ days: 7 })
			.toLocaleString(DateTime.DATE_FULL)}\n\n`;
		let output = headline + "| Subject | Day | Start | End | Done |\n";
		output += "| --- | --- | --- | --- | --- |\n";
		for (let event of events.value) {
			output += `| ${event.subject} | ${DateTime.fromISO(
				event.start.dateTime
			).toLocaleString(DateTime.DATE_SHORT)} | ${DateTime.fromISO(
				event.start.dateTime
			).toLocaleString(DateTime.TIME_24_SIMPLE)} | ${DateTime.fromISO(
				event.end.dateTime
			).toLocaleString(DateTime.TIME_24_SIMPLE)} |  |\n`;
		}
		return output;ß
	}

	public async getUserEventsForToday() {
		const today = new Date();
		today.setUTCHours(0, 0, 0, 0);

		const tomorrow = new Date();
		tomorrow.setDate(today.getDate() + 1);
		tomorrow.setUTCHours(0, 0, 0, 0);

		const events = await this.graphClient
			.api("/me/calendarview")
			.query({
				startDateTime: today.toISOString(),
				endDateTime: tomorrow.toISOString(),
			})
			.get();

		return events;
	}

	public async getUserEventsForWeek() {
		const today = new Date();
		today.setUTCHours(0, 0, 0, 0);

		const week = new Date();
		week.setDate(today.getDate() + 7);
		week.setUTCHours(0, 0, 0, 0);

		const events = await this.graphClient
			.api("/me/calendarview")
			.query({
				startDateTime: today.toISOString(),
				endDateTime: week.toISOString(),
			})
			.get();

		return events;
	}
}
