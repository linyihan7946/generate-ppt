import axios from 'axios';

export class ImageService {
    private apiKey: string | undefined;
    private baseUrl: string;

    constructor() {
        this.apiKey = process.env.IMAGE_API_KEY;
        this.baseUrl = process.env.IMAGE_API_BASE_URL || 'https://www.aigenimage.cn';
        console.log('ImageService initialized with baseUrl:', this.baseUrl);
    }

    async generateImage(prompt: string): Promise<string | null> {
        try {
            const headers: any = {
                'Content-Type': 'application/json',
            };

            if (this.apiKey) {
                headers['Authorization'] = `Bearer ${this.apiKey}`;
            }

            const url = `${this.baseUrl}/api/image/direct-edit`;
            console.log('Making request to:', url);

            const response = await axios.post(
                url,
                {
                    prompt: `Professional presentation illustration: ${prompt}`,
                    images: [],
                    model: 'gemini-3.1-flash-image-preview',
                    aspect_ratio: '16:9',
                    resolution: '2K'
                },
                {
                    headers,
                    timeout: 120000,
                    proxy: false
                }
            );

            console.log('API Response:', JSON.stringify(response.data, null, 2));

            if (response.data?.success && response.data?.data?.data?.[0]) {
                const result = response.data.data.data[0];
                return typeof result === 'string' ? result : result.url;
            }

            console.error('Invalid response format:', response.data);
            return null;
        } catch (error: any) {
            console.error('Error generating image:', error instanceof Error ? error.message : error);
            if (error.response) {
                console.error('Error response data:', JSON.stringify(error.response.data, null, 2));
            }
            return null;
        }
    }
}