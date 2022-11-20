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
	Setting,
} from "obsidian";

// Remember to rename these classes and interfaces!

export interface MyPluginSettings {
	folderName: string;
}

const DEFAULT_SETTINGS: MyPluginSettings = {
	folderName: "",
};

export default class MyPlugin extends Plugin {
	settings: MyPluginSettings;
	authProvider: DeviceCodeAuthProvider;
	graphClient: Client;
	graphExplorer: GraphExplorer;

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

		this.graphExplorer = new GraphExplorer(this.graphClient, this.settings);

		// This adds an editor command that can perform some operation on the current editor instance
		this.addCommand({
			id: "add-graph-day-events-table",
			name: "O365 Today Events",
			editorCallback: async (editor: Editor, view: MarkdownView) => {
				const eventsTable = await this.graphExplorer.obsidianRenderTodaysEvent();

				editor.replaceSelection(eventsTable);
			},
		});

		this.addCommand({
			id: "add-graph-week-events-table",
			name: "O365 Week Events",
			editorCallback: async (editor: Editor, view: MarkdownView) => {
				const eventsTable = await this.graphExplorer.obsidianRenderWeekEvents();

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

		containerEl.createEl("h3", { text: "Settings" });

		new Setting(containerEl)
			.setName("Meetings Folder")
			.setDesc("Set the folder where the meeting notes will be saved")
			.addText((text) =>
				text
					.setPlaceholder("Enter folder name")
					.setValue(this.plugin.settings.folderName)
					.onChange(async (value) => {
						this.plugin.settings.folderName = value;
						await this.plugin.saveSettings();
					})
			);

		const c = containerEl.createEl("h2", {
			text: this.userEmail
				? `Logged in as: ${this.userEmail}`
				: "Not logged in",
		});

		c.style.marginTop = "20px";

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
