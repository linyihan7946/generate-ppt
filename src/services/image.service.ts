import axios from 'axios';

export class ImageService {
    private openaiApiKey: string | undefined;

    constructor() {
        this.openaiApiKey = process.env.OPENAI_API_KEY;
    }

    async generateImage(prompt: string): Promise<string | null> {
        if (!this.openaiApiKey) {
            console.log('No OpenAI API key provided, skipping AI image generation.');
            return null;
        }

        try {
            const response = await axios.post(
                'https://api.openai.com/v1/images/generations',
                {
                    prompt: `Professional presentation illustration for: ${prompt}`,
                    n: 1,
                    size: '1024x1024',
                },
                {
                    headers: {
                        Authorization: `Bearer ${this.openaiApiKey}`,
                    },
                }
            );

            return response.data.data[0].url;
        } catch (error) {
            console.error('Error generating image:', error);
            return null;
        }
    }
}
