import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as SpeechSDK from "microsoft-cognitiveservices-speech-sdk";

export class MainControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _value: string;
    private _silenceTimer: number | null = null;
    private _activeButtonText = "Click here and start talking 🎙️";
    private _inactiveButtonText = "I'm listening 👂...";

    private _context: ComponentFramework.Context<IInputs>;
    private _notifyOutputChanged: () => void;
    private _container: HTMLDivElement;
    public _inputText: HTMLTextAreaElement;
    private _speakNowButton: HTMLButtonElement;  

    private _recognizer: SpeechSDK.SpeechRecognizer | null = null;
    ///private _speechConfig: SpeechSDK.SpeechConfig;
    //private _audioConfig: SpeechSDK.AudioConfig;

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
           this._context = context;
        this._notifyOutputChanged = notifyOutputChanged;

        this._container = document.createElement("div");
        this._container.classList.add("MainControl");

        this._inputText = document.createElement("textarea");
        this._inputText.setAttribute("class", "input-text");
        this._inputText.value = this._value || "";
        this._inputText.addEventListener("input", this._refreshData);

        this._speakNowButton = document.createElement("button");
        this._speakNowButton.setAttribute("type", "button");
        this._speakNowButton.setAttribute("class", "record-button");
        this._speakNowButton.innerText = this._activeButtonText;

        // Toggle logic
        this._speakNowButton.addEventListener("click", () => {
            if (this._speakNowButton.innerText === this._activeButtonText) {
                this.StartSpeechRecognition();
            } else {
                this.StopSpeechRecognition();
            }
        });

        this._container.append(this._speakNowButton, this._inputText);
        container.append(this._container);
    }

    private _refreshData = (): void => {
        this._value = this._inputText.value;
        this._notifyOutputChanged();
    };
    private startSilenceTimer(): void {
        if (this._silenceTimer) window.clearTimeout(this._silenceTimer);
        this._silenceTimer = window.setTimeout(() => this.StopSpeechRecognition(), 5000);
    }

private async getSecretValue(envVarName: string): Promise<string> {
    const query = `?$filter=schemaname eq '${envVarName}'&$expand=environmentvariabledefinition_environmentvariablevalue($select=value)`;
    const result = await this._context.webAPI.retrieveMultipleRecords("environmentvariabledefinition", query);
    
    if (result.entities && result.entities.length > 0) {
        // Entity type and access via string key to satisfy ESLint
        const definition = result.entities[0] as ComponentFramework.WebApi.Entity;
        
        const relationKey = "environmentvariabledefinition_environmentvariablevalue";
        const values = definition[relationKey];
        
        if (Array.isArray(values) && values.length > 0) {
            return values[0].value || "";
        }
        
        // Fallback to default value if current value isn't set
        return (definition["defaultvalue"] as string) || "";
    }
    return "";
}






private async StartSpeechRecognition(): Promise<void> {
    try{
    //  credentials from D365 Environment Variables
            const key = await this.getSecretValue(this._context.parameters.SpeechKeyVar.raw || "");
            const region = await this.getSecretValue(this._context.parameters.SpeechRegionVar.raw || "");

            if (!key || !region) {
                alert("Speech Key or Region missing in Environment Variables.");
                return;
            }


             // 2. Setup Config
            const speechConfig = SpeechSDK.SpeechConfig.fromSubscription(key, region);
            speechConfig.speechRecognitionLanguage = "en-US";
            const audioConfig = SpeechSDK.AudioConfig.fromDefaultMicrophoneInput();

    this._recognizer = new SpeechSDK.SpeechRecognizer(speechConfig, audioConfig);
            this._speakNowButton.innerText = this._inactiveButtonText;

    // Logic: Clear timer while speaking
    this._recognizer.recognizing = (s, e) => {
        if (this._silenceTimer) window.clearTimeout(this._silenceTimer);
        this._inputText.value = (this._value + " " + e.result.text).trim();
    };

    // 2. START timer only after a phrase is finalized
    this._recognizer.recognized = (s, e) => {
        if (e.result.reason === SpeechSDK.ResultReason.RecognizedSpeech) {
            this._value += " " + e.result.text;
            this._inputText.value = this._value.trim();
            this._refreshData();
            
            // User finished a phrase; start the 5s "wait for more" countdown
            this.startSilenceTimer(); 
        }
    };

    this._recognizer.sessionStopped = () => this.ResetUI();
    this._recognizer.startContinuousRecognitionAsync();
}
catch (error) {
            console.error("Secure Auth Failed:", error);
        }
}


    private ResetUI(): void {
        if (this._silenceTimer) window.clearTimeout(this._silenceTimer);
        this._speakNowButton.innerText = this._activeButtonText;
    }

    private StopSpeechRecognition(): void {
        if (this._recognizer) {
            this._recognizer.stopContinuousRecognitionAsync(() => {
                this._recognizer?.close();
                this._recognizer = null;
                this.ResetUI();
            });
        }
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._inputText.value = context.parameters.sampleProperty.raw || "";
    }

    public getOutputs(): IOutputs { return { sampleProperty: this._value }; }

    public destroy(): void { this.StopSpeechRecognition(); }
}
