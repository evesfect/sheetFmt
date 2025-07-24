package mapping

import (
	"context"
	"fmt"
	"os"
	"strings"
	"time"

	"github.com/google/generative-ai-go/genai"
	"google.golang.org/api/option"
)

// AIMapping represents an AI-suggested mapping with confidence
type AIMapping struct {
	ScannedColumn string  `json:"scanned_column"`
	TargetColumn  string  `json:"target_column"`
	Confidence    float64 `json:"confidence"`
}

// AIConfig holds configuration for AI mapping
type AIConfig struct {
	APIKey    string
	Model     string
	BatchSize int
}

// AIMapper handles AI-powered column mapping
type AIMapper struct {
	client *genai.Client
	model  *genai.GenerativeModel
	config AIConfig
}

// NewAIMapper creates a new AI mapper instance
func NewAIMapper(apiKey string) (*AIMapper, error) {
	if apiKey == "" {
		return nil, fmt.Errorf("gemini API key is required")
	}

	ctx := context.Background()
	client, err := genai.NewClient(ctx, option.WithAPIKey(apiKey))
	if err != nil {
		return nil, fmt.Errorf("failed to create Gemini client: %v", err)
	}

	model := client.GenerativeModel("gemini-1.5-flash")
	model.SetTemperature(0.1) // Low temperature for consistent results

	return &AIMapper{
		client: client,
		model:  model,
		config: AIConfig{
			APIKey:    apiKey,
			Model:     "gemini-1.5-flash",
			BatchSize: 10,
		},
	}, nil
}

// Close cleans up the AI mapper resources
func (ai *AIMapper) Close() error {
	if ai.client != nil {
		return ai.client.Close()
	}
	return nil
}

func (ai *AIMapper) GenerateColumnMappings(scannedColumns, targetColumns []string) ([]AIMapping, error) {
	if len(scannedColumns) == 0 || len(targetColumns) == 0 {
		return nil, fmt.Errorf("both scanned and target columns must be provided")
	}

	debugLog("Building prompt for %d scanned columns and %d target columns", len(scannedColumns), len(targetColumns))

	ctx, cancel := context.WithTimeout(context.Background(), 200*time.Second)
	defer cancel()

	prompt := ai.buildMappingPrompt(scannedColumns, targetColumns)
	debugLog("Generated prompt length: %d characters", len(prompt))

	debugLog("Sending request to Gemini API...")
	resp, err := ai.model.GenerateContent(ctx, genai.Text(prompt))
	if err != nil {
		debugLog("Gemini API request failed: %v", err)
		return nil, fmt.Errorf("failed to generate AI response: %v", err)
	}

	if len(resp.Candidates) == 0 || len(resp.Candidates[0].Content.Parts) == 0 {
		debugLog("No response candidates received from Gemini API")
		return nil, fmt.Errorf("no response generated from AI")
	}

	// Extract text from response
	var responseText string
	for _, part := range resp.Candidates[0].Content.Parts {
		if textPart, ok := part.(genai.Text); ok {
			responseText += string(textPart)
		}
	}

	debugLog("Received response from Gemini API, length: %d characters", len(responseText))
	debugLog("Raw AI response:\n%s", responseText)

	mappings, err := ai.parseMappingResponse(responseText)
	if err != nil {
		debugLog("Failed to parse AI response: %v", err)
		return nil, fmt.Errorf("failed to parse AI response: %v", err)
	}

	debugLog("Successfully parsed %d mappings from AI response", len(mappings))
	return mappings, nil
}

// buildMappingPrompt creates a prompt for the AI to generate column mappings
func (ai *AIMapper) buildMappingPrompt(scannedColumns, targetColumns []string) string {
	prompt := `You are an expert data analyst helping to map column names from various Excel files to a standardized target format.

TASK: Map each scanned column to the most appropriate target column, or mark as "NO_MATCH" if uncertain.

SCANNED COLUMNS (from various Excel files):
`
	for _, col := range scannedColumns {
		prompt += fmt.Sprintf("- %s\n", col)
	}

	prompt += `
TARGET COLUMNS (standardized format):
`
	for _, col := range targetColumns {
		prompt += fmt.Sprintf("- %s\n", col)
	}

	prompt += `
INSTRUCTIONS:
1. Only suggest mappings you are confident about (>80% certainty)
2. Consider semantic meaning, not just text similarity
3. Map each scanned column to AT MOST ONE target column
4. If uncertain or no clear match exists, use "NO_MATCH"

OUTPUT FORMAT (one line per scanned column):
ScannedColumn|TargetColumn|Confidence

EXAMPLES:
Customer Name|Name|0.95
Cust_ID|ID|0.90
Phone Number|Phone|0.95
Random_Data|NO_MATCH|0.00

Now provide mappings for the scanned columns:`

	return prompt
}

// parseMappingResponse parses the AI response into structured mappings
func (ai *AIMapper) parseMappingResponse(response string) ([]AIMapping, error) {
	var mappings []AIMapping
	lines := strings.Split(strings.TrimSpace(response), "\n")

	for _, line := range lines {
		line = strings.TrimSpace(line)
		if line == "" || strings.HasPrefix(line, "ScannedColumn|") {
			continue
		}

		parts := strings.Split(line, "|")
		if len(parts) != 3 {
			continue
		}

		scannedCol := strings.TrimSpace(parts[0])
		targetCol := strings.TrimSpace(parts[1])
		confidenceStr := strings.TrimSpace(parts[2])

		// Parse confidence
		var confidence float64
		if _, err := fmt.Sscanf(confidenceStr, "%f", &confidence); err != nil {
			confidence = 0.0
		}

		// Skip if NO_MATCH or low confidence
		if targetCol == "NO_MATCH" || confidence < 0.8 {
			continue
		}

		mappings = append(mappings, AIMapping{
			ScannedColumn: scannedCol,
			TargetColumn:  targetCol,
			Confidence:    confidence,
		})
	}

	return mappings, nil
}

// GetGeminiAPIKey gets the API key from environment variable
func GetGeminiAPIKey() string {
	return os.Getenv("GEMINI_API_KEY")
}
