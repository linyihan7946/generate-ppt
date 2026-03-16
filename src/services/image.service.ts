import axios from 'axios';

export class ImageService {
    private apiKey: string | undefined;
    private baseUrl: string;

    constructor() {
        this.apiKey = process.env.IMAGE_API_KEY;
        this.baseUrl = process.env.IMAGE_API_BASE_URL || 'https://www.aigenimage.cn';
    }

    async generateImage(prompt: string): Promise<string | null> {
        try {
            const headers: any = {
                'Content-Type': 'application/json',
            };

            if (this.apiKey) {
                headers['Authorization'] = `Bearer ${this.apiKey}`;
            }

            const response = await axios.post(
                `${this.baseUrl}/api/image/direct-edit`,
                {
                    prompt: `Professional presentation illustration: ${prompt}`,
                    images: [],
                    model: 'gemini-3-pro-image-preview',
                    aspect_ratio: '16:9',
                    resolution: '2K'
                },
                {
                    headers,
                    timeout: 60000
                }
            );

            if (response.data?.success && response.data?.data?.data?.[0]) {
                return response.data.data.data[0];
            }

            console.error('Invalid response format:', response.data);
            return null;
        } catch (error) {
            console.error('Error generating image:', error instanceof Error ? error.message : error);
            return null;
        }
    }
}