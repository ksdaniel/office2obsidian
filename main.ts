import { Client } from "@microsoft/microsoft-graph-client";
import DeviceCodeAuthProvider from "DeviceCodeAuthProvider";
import GraphExplorer from "GraphExporer";
import {
	App,
	ButtonComponent,
	Editor,
	MarkdownView,
	Plugin,
	PluginSettingTab,
} from "obsidian";

// Remember to rename these classes and interfaces!

interface MyPluginSettings {
	mySetting: string;
}

const DEFAULT_SETTINGS: MyPluginSettings = {
	mySetting: "default",
};

export default class MyPlugin extends Plugin {
	settings: MyPluginSettings;
	authProvider: DeviceCodeAuthProvider;
	graphClient: Client;

	async onload() {
		await this.loadSettings();

		this.authProvider = new DeviceCodeAuthProvider(
			"150e3906-8875-4502-8933-8d8fad4f26d2",
			this.app
		);

		this.graphClient = Client.initWithMiddleware({
			authProvider: this.authProvider,
			defaultVersion: "beta",
		});

		// This adds an editor command that can perform some operation on the current editor instance
		this.addCommand({
			id: "add-graph-day-events-table",
			name: "O365 Today Events",
			editorCallback: async (editor: Editor, view: MarkdownView) => {
				const eventsTable = await new GraphExplorer(
					this.graphClient
				).obsidianRenderTodaysEvent();

				editor.replaceSelection(eventsTable);
			},
		});

		this.addCommand({
			id: "add-graph-week-events-table",
			name: "O365 Week Events",
			editorCallback: async (editor: Editor, view: MarkdownView) => {
				const eventsTable = await new GraphExplorer(
					this.graphClient
				).obsidianRenderWeekEvents();

				editor.replaceSelection(eventsTable);
			},
		});

		// This adds a settings tab so the user can configure various aspects of the plugin
		this.addSettingTab(new SampleSettingTab(this.app, this));

		// If the plugin hooks up any global DOM events (on parts of the app that doesn't belong to this plugin)
		// Using this function will automatically remove the event listener when this plugin is disabled.
		this.registerDomEvent(document, "click", (evt: MouseEvent) => {
			console.log("click", evt);
		});

		// When registering intervals, this function will automatically clear the interval when the plugin is disabled.
		this.registerInterval(
			window.setInterval(() => console.log("setInterval"), 5 * 60 * 1000)
		);
	}

	onunload() {}

	async loadSettings() {
		this.settings = Object.assign(
			{},
			DEFAULT_SETTINGS,
			await this.loadData()
		);
	}

	async saveSettings() {
		await this.saveData(this.settings);
	}
}

class SampleSettingTab extends PluginSettingTab {
	plugin: MyPlugin;
	userName: string;
	userEmail: string;
	isLoggedIn: boolean;

	constructor(app: App, plugin: MyPlugin) {
		super(app, plugin);
		this.plugin = plugin;

		console.log("sample setting tab constructor");

		this.plugin.authProvider.onLogin = async () => {
			console.log("onLogin");

			this.plugin.authProvider.getloggedInUser().then((user) => {
				if (user && user.name) {
					this.userName = user?.name;
					this.userEmail = user?.username;
					this.isLoggedIn = true;
				}
			});

			this.display();
		};

		this.plugin.authProvider.onLogout = async () => {
			console.log("onLogout");
			this.isLoggedIn = false;
			this.userEmail = "";
			this.userName = "";
			this.display();
		};

		this.plugin.authProvider.getloggedInUser().then((user) => {
			if (user && user.name) {
				this.userName = user?.name;
				this.userEmail = user?.username;
				this.isLoggedIn = true;
			}
		});
	}

	display(): void {
		const { containerEl } = this;

		containerEl.empty();

		containerEl.createEl("h2", { text: "O365 to Obsidian events import" });

		containerEl.createEl("h3", {
			text: this.userEmail
				? `Logged in as: ${this.userEmail}`
				: "Not logged in",
		});

		new ButtonComponent(containerEl)
			.setButtonText(this.isLoggedIn ? "Logout" : "Login")
			.onClick(async (evt) => {
				if (this.isLoggedIn) {
					await this.plugin.authProvider.logout();
					this.display();
				} else {
					await this.plugin.authProvider.getAccessToken();
					this.display();
				}
			});
	}
}
