import { ImageService } from './src/services/image.service';
import * as dotenv from 'dotenv';
import * as path from 'path';

// Load environment variables from .env file
dotenv.config({ path: path.resolve(__dirname, '.env') });

async function testImageGeneration() {
    console.log('Testing Image Generation API...');
    
    // Check if API key is present
    if (!process.env.IMAGE_API_KEY) {
        console.warn('Warning: IMAGE_API_KEY is not set in .env file.');
    } else {
        console.log('IMAGE_API_KEY is set.');
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
            console.log('✅ Image generation successful!');
            console.log('Image URL:', imageUrl);
            console.log(`Time taken: ${(endTime - startTime) / 1000} seconds`);
        } else {
            console.error('❌ Image generation failed. No URL returned.');
        }
    } catch (error) {
        console.error('❌ An error occurred during image generation:', error);
    }
}

testImageGeneration();
