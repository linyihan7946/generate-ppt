import { ImageService } from '../src/services/image.service';
import * as dotenv from 'dotenv';
import * as path from 'path';

// Load environment variables from .env file
dotenv.config({ path: path.resolve(__dirname, '../.env') });

async function testImageGeneration() {
    console.log('Testing Image Generation API...');

    const hasOpenAiKey = Boolean(process.env.OPENAI_API_KEY);
    const hasLegacyKey = Boolean(process.env.IMAGE_API_KEY);

    if (!hasOpenAiKey && !hasLegacyKey) {
        console.warn('Warning: neither OPENAI_API_KEY nor IMAGE_API_KEY is set.');
    } else {
        console.log(`Configured provider: ${hasOpenAiKey ? 'OPENAI /images/edits' : 'legacy direct-edit'}`);
    }

    const imageService = new ImageService();
    // Use a simple prompt for testing
    const prompt = 'A serene mountain landscape with a lake at sunrise';

    try {
        console.log(`Generating image with prompt: "${prompt}"`);
        const startTime = Date.now();

        // The service adds "Professional presentation illustration: " prefix automatically
        const imageUrl = await imageService.generateImage(prompt);

        const endTime = Date.now();

        if (imageUrl) {
            console.log('[OK] Image generation successful!');
            console.log('Image preview:', imageUrl.slice(0, 80) + '...');
            console.log('Payload length:', imageUrl.length);
            console.log(`Time taken: ${(endTime - startTime) / 1000} seconds`);
        } else {
            console.error('[FAIL] Image generation failed. No URL returned.');
        }
    } catch (error) {
        console.error('[FAIL] An error occurred during image generation:', error);
    }
}

testImageGeneration();
