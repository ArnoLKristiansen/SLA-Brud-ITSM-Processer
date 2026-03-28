import { useState } from "react";
import { PowerApps } from "@microsoft/power-apps";
import "./App.css";

export default function App() {
    const [text, setText] = useState("");
    const [loading, setLoading] = useState(false);

    const saveNote = async () => {
        if (!text.trim()) {
            alert("Venligst skriv en note");
            return;
        }

        setLoading(true);
        try {
            const pa = new PowerApps();

            await pa.dataverse.createRecord("bd_slanotes", {
                bd_notetext: text
            });

            alert("Gemt i Dataverse ✅");
            setText("");
        } catch (error) {
            alert("Fejl ved gemning: " + (error instanceof Error ? error.message : "Ukendt fejl"));
        } finally {
            setLoading(false);
        }
    };

    return (
        <div className="app-container">
            <div className="app-card">
                <div className="app-header">
                    <h1>Velkommen til SLA Rapportering</h1>
                    <p className="subtitle">ITSM Processes</p>
                </div>

                <div className="form-section">
                    <label htmlFor="noteInput" className="form-label">
                        Skriv en SLA note
                    </label>
                    
                    <textarea
                        id="noteInput"
                        value={text}
                        onChange={(e) => setText(e.target.value)}
                        placeholder="Indtast din note her..."
                        disabled={loading}
                        className="form-textarea"
                        rows={6}
                    />

                    <button 
                        onClick={saveNote} 
                        disabled={loading}
                        className="submit-button"
                    >
                        {loading ? "Gemmer..." : "Gem i Dataverse"}
                    </button>
                </div>

                <div className="info-section">
                    <p className="info-text">
                        Dine noter gemmes automatisk i Dataverse
                    </p>
                </div>
            </div>
        </div>
    );
}