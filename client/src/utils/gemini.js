const GEMINI_API_KEY = "AIzaSyBk7paPnpP523j2LKgRegRZFCCMOKaHmrM";
const BASE_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent";

export async function fetchGeminiData(prompt) {
    try {
        const response = await fetch(`${BASE_URL}?key=${GEMINI_API_KEY}`, {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
            },
            body: JSON.stringify({
                contents: [
                    {
                        parts: [
                            {
                                text: `${prompt}\n\nReturn ONLY a valid JSON object. No markdown, no triple backticks, just the raw JSON.`,
                            },
                        ],
                    },
                ],
                generationConfig: {
                    response_mime_type: "application/json",
                },
            }),
        });

        if (!response.ok) {
            const errorData = await response.json();
            console.error("Gemini API Error:", errorData);
            throw new Error(`Gemini API Error: ${response.statusText}`);
        }

        const data = await response.json();
        const textContent = data.candidates?.[0]?.content?.parts?.[0]?.text;

        if (!textContent) {
            throw new Error("No content returned from Gemini");
        }

        return JSON.parse(textContent);
    } catch (error) {
        console.error("Error fetching data from Gemini:", error);
        return null;
    }
}
