import { Modal, App } from "obsidian";

export class LoginModal extends Modal {

    public message = "Hello World";

	constructor(app: App) {    
		super(app);
	}

    public setMessage(message: string) {
        this.message = message;
    }

	onOpen() {
		const {contentEl} = this;
        const messageDiv = document.createElement("div")
        messageDiv.innerHTML = this.message;
        messageDiv.style.userSelect = "text";
        contentEl.appendChild(messageDiv);
	}

	onClose() {
		const {contentEl} = this;
		contentEl.empty();
	}
}