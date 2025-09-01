# Sales Proposal Bot

A Slack bot that automatically adds template slides to PowerPoint proposals based on location requests.

## Features

- Natural language interaction through Slack
- Automatic location detection from user messages
- Adds financial proposal slide as the second-to-last slide in presentations
- Dynamic financial calculations with VAT (5% of subtotal)
- Multiple duration/pricing options support
- Automatic slide scaling to match presentation dimensions
- Supports both direct messages and slash commands
- REST API for programmatic access

## Setup

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Create a `.env` file with your credentials:
```bash
cp .env.example .env
# Edit .env with your actual credentials
```

3. Ensure you have the following files in the directory:
   - `image.png` - Header image for the financial proposal slide
   - Location-specific presentations:
     - `1. Desirable by Location - The Landmark Series copy 2.pptx`
     - `Jawhara.pptx`
     - `The Gateway.pptx`
     - `The Oryx.pptx`
     - `The Triple Crown.pptx`

4. Run the bot:
```bash
python main.py
```

## Usage

### Creating Proposals

Message the bot naturally:
- "I need a proposal for landmark starting January 1st, 2 weeks at 1.5M"
- "Gateway proposal for Feb 15th, options for 2 weeks (2M) and 4 weeks (3.5M)"
- "Create oryx proposal, March 1st start, 6 weeks duration, 4.5M net rate"

The bot will:
1. Ask for any missing information (location, start date, durations, net rates)
2. Validate that the number of durations matches the number of net rates
3. Create a financial proposal slide with:
   - Location details
   - Start date
   - Multiple duration/pricing options in columns
   - Automatic VAT calculation (5%)
   - Total amounts for each option
4. Insert the slide as second-to-last in the presentation
5. Send the completed PowerPoint file via Slack

### Available Locations

- **Landmark** (or "The Landmark")
- **Jawhara**
- **Gateway** (or "The Gateway")
- **Oryx** (or "The Oryx")
- **Triple Crown**

### API Endpoints

- `POST /api/proposal` - Get a proposal programmatically
  ```json
  {
    "location": "landmark"
  }
  ```

- `GET /api/locations` - Get list of available locations

- `GET /health` - Health check endpoint

## How It Works

1. User requests a proposal with location, start date, durations, and net rates
2. Bot uses OpenAI to understand the request and extract details
3. Bot retrieves the corresponding PowerPoint file for the location
4. Financial proposal slide is created with:
   - Dynamic VAT calculation (5% of subtotal)
   - Multiple duration/pricing options displayed in columns
   - All elements scaled to match presentation dimensions
   - Professional formatting with colors and borders
5. Slide is inserted as the second-to-last slide in the presentation
6. Modified presentation is sent back to the user via Slack

## Configuration

The bot uses environment variables for configuration:

- `SLACK_BOT_TOKEN` - Your Slack bot token
- `SLACK_SIGNING_SECRET` - Slack signing secret for verification
- `OPENAI_API_KEY` - OpenAI API key for natural language processing

## Development

The bot is built with:
- FastAPI for the web framework
- Slack SDK for Slack integration
- OpenAI API for natural language understanding
- python-pptx for PowerPoint manipulation

## Troubleshooting

1. **Bot not responding**: Check that all environment variables are set correctly
2. **File not found errors**: Ensure all PowerPoint files are in the same directory as the script
3. **Permission errors**: Make sure the bot has been added to the Slack channel

## Security Notes

- The bot verifies all Slack requests using the signing secret
- Environment variables should never be committed to version control
- Use `.env` files for local development only