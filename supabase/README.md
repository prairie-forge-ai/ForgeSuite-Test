# Ada - Prairie Forge AI Assistant

Ada is Prairie Forge's AI-powered financial analyst, named after Ada Lovelace.
Powered by ChatGPT (OpenAI GPT-4 Turbo).

## Cost Estimates

| Usage Level | Questions/Day | Monthly Cost |
|-------------|---------------|--------------|
| Light | 10 | ~$6 |
| Moderate | 50 | ~$30 |
| Heavy | 200 | ~$120 |

*Based on GPT-4 Turbo pricing: ~$0.02-0.05 per question*

## Setup Instructions

### 1. Create Supabase Project

If you don't have one:
```bash
# Install Supabase CLI
npm install -g supabase

# Login to Supabase
supabase login

# Create a new project at https://supabase.com/dashboard
```

### 2. Link Your Project

```bash
cd Customer-ArchCollins-Foundry
supabase link --project-ref YOUR_PROJECT_REF
```

### 3. Set Your OpenAI API Key

```bash
supabase secrets set OPENAI_API_KEY=sk-your-openai-key-here
```

Get your API key at: https://platform.openai.com/api-keys

### 4. Deploy Ada

```bash
supabase functions deploy copilot --no-verify-jwt
```

The `--no-verify-jwt` flag allows the Excel add-in to call Ada without authentication.

### 5. Update the Frontend

In `payroll-recorder/src/workflow.js`, update:

```javascript
const COPILOT_API_ENDPOINT = "https://YOUR_PROJECT.supabase.co/functions/v1/copilot";

bindCopilotCard(container, { 
    id: "expense-review-copilot",
    apiEndpoint: COPILOT_API_ENDPOINT,  // ← Uncomment this line
    // ... rest of config
});
```

### 6. Test It!

```bash
curl -X POST https://YOUR_PROJECT.supabase.co/functions/v1/copilot \
  -H "Content-Type: application/json" \
  -d '{
    "prompt": "What should I focus on this payroll period?",
    "context": {
      "period": "Nov 2025",
      "summary": { "total": 253625, "employeeCount": 38 }
    }
  }'
```

## Configuration

### Change the AI Model

In `supabase/functions/copilot/index.ts`:

```typescript
// For cheaper responses (less accurate):
const DEFAULT_MODEL = "gpt-3.5-turbo";

// For best quality (current default):
const DEFAULT_MODEL = "gpt-4-turbo-preview";
```

### Customize Ada's Personality

Edit the `defaultSystemPrompt` in the Edge Function to change how Ada responds.

## Switching to Claude (Future)

When ready to switch to Claude:

1. Get Anthropic API key
2. Update the Edge Function to call Anthropic's API
3. Change "Powered by ChatGPT" to "Powered by Claude"

The interface stays the same — just swap the backend!

## Files

```
supabase/
├── config.toml              # Supabase project config
├── functions/
│   └── copilot/
│       └── index.ts         # Ada's backend (Edge Function)
└── README.md                # This file
```

## Troubleshooting

**"Ada is not configured yet"**
- Make sure `OPENAI_API_KEY` secret is set
- Run: `supabase secrets list` to verify

**"Ada is thinking hard right now"**
- OpenAI rate limiting — wait a moment and retry

**Ada's responses seem generic**
- Make sure context is being passed from the spreadsheet
- Check browser console for context provider errors
