package mapping

import (
	"context"
	"fmt"
	"os"
	"sheetFmt/internal/logger"
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

// AIMapper handles AI-powered column mapping
type AIMapper struct {
	client *genai.Client
	model  *genai.GenerativeModel
}

// NewAIMapper creates a new AI mapper instance
func NewAIMapper(apiKey string) (*AIMapper, error) {
	if apiKey == "" {
		return nil, fmt.Errorf("gemini API key is required")
	}

	logger.Info("Initializing AI mapper with Gemini API")
	logger.Debug("API key length", "length", len(apiKey))

	ctx := context.Background()
	client, err := genai.NewClient(ctx, option.WithAPIKey(apiKey))
	if err != nil {
		logger.Error("Failed to create Gemini client", "error", err)
		return nil, fmt.Errorf("failed to create Gemini client: %v", err)
	}

	// Use the latest Gemini 2.0 Flash model
	modelName := "gemini-2.0-flash-exp"
	model := client.GenerativeModel(modelName)
	model.SetTemperature(0.1) // Low temperature for consistent results

	logger.Info("AI mapper initialized successfully", "model", modelName, "temperature", 0.1)

	return &AIMapper{
		client: client,
		model:  model,
	}, nil
}

// Close cleans up the AI mapper resources
func (ai *AIMapper) Close() error {
	if ai.client != nil {
		logger.Debug("Closing AI mapper client")
		return ai.client.Close()
	}
	return nil
}

func (ai *AIMapper) GenerateColumnMappings(scannedColumns, targetColumns []string) ([]AIMapping, error) {
	if len(scannedColumns) == 0 || len(targetColumns) == 0 {
		return nil, fmt.Errorf("both scanned and target columns must be provided")
	}

	logger.Info("=== AI MAPPING REQUEST START ===")
	logger.Info("Generating AI column mappings",
		"scanned_count", len(scannedColumns),
		"target_count", len(targetColumns))

	// Log the actual columns being processed
	logger.Debug("=== SCANNED COLUMNS ===")
	for i, col := range scannedColumns {
		logger.Debug("Scanned column", "index", i+1, "name", col)
	}

	logger.Debug("=== TARGET COLUMNS ===")
	for i, col := range targetColumns {
		logger.Debug("Target column", "index", i+1, "name", col)
	}

	// 50 columns is fine, only chunk if we have 100+ columns
	if len(scannedColumns) > 100 {
		logger.Info("Very large request detected, processing in chunks",
			"total_columns", len(scannedColumns))
		return ai.generateMappingsInChunks(scannedColumns, targetColumns, 50)
	}

	return ai.generateSingleBatch(scannedColumns, targetColumns)
}

func (ai *AIMapper) generateMappingsInChunks(scannedColumns, targetColumns []string, chunkSize int) ([]AIMapping, error) {
	var allMappings []AIMapping
	totalChunks := (len(scannedColumns) + chunkSize - 1) / chunkSize

	logger.Info("Processing in chunks", "total_chunks", totalChunks, "chunk_size", chunkSize)

	// Process in chunks
	for i := 0; i < len(scannedColumns); i += chunkSize {
		end := i + chunkSize
		if end > len(scannedColumns) {
			end = len(scannedColumns)
		}

		chunk := scannedColumns[i:end]
		chunkNum := (i / chunkSize) + 1

		logger.Info("Processing chunk",
			"chunk", chunkNum,
			"total_chunks", totalChunks,
			"range", fmt.Sprintf("%d-%d", i+1, end),
			"size", len(chunk))

		chunkMappings, err := ai.generateSingleBatch(chunk, targetColumns)
		if err != nil {
			logger.Error("Failed to process chunk", "chunk", chunkNum, "error", err)
			// Continue with other chunks instead of failing completely
			continue
		}

		logger.Info("Chunk processed successfully", "chunk", chunkNum, "mappings_found", len(chunkMappings))
		allMappings = append(allMappings, chunkMappings...)

		if chunkNum < totalChunks {
			logger.Debug("Waiting between chunks to avoid rate limiting", "delay", "2s")
			time.Sleep(2 * time.Second)
		}
	}

	logger.Info("All chunks processed", "total_mappings", len(allMappings))
	return allMappings, nil
}

func (ai *AIMapper) generateSingleBatch(scannedColumns, targetColumns []string) ([]AIMapping, error) {
	logger.Info("Generating single batch AI mappings")

	// Build the prompt
	startTime := time.Now()
	prompt := ai.buildMappingPrompt(scannedColumns, targetColumns)
	promptBuildTime := time.Since(startTime)

	logger.Info("Prompt generated",
		"length", len(prompt),
		"build_time", promptBuildTime)

	// Log the full prompt for debugging
	logger.Info("FULL PROMPT SENT TO AI", "content", prompt)

	// Increased timeout since 30s was too short
	timeout := 60 * time.Second
	ctx, cancel := context.WithTimeout(context.Background(), timeout)
	defer cancel()

	logger.Info("Sending request to Gemini API", "timeout", timeout)

	// Create a channel to handle the response
	type apiResult struct {
		resp *genai.GenerateContentResponse
		err  error
	}

	resultChan := make(chan apiResult, 1)
	apiStartTime := time.Now()

	// Make the API call in a goroutine
	go func() {
		logger.Info("API call started")
		resp, err := ai.model.GenerateContent(ctx, genai.Text(prompt))
		apiCallTime := time.Since(apiStartTime)
		logger.Info("API call completed", "duration", apiCallTime, "has_error", err != nil)
		resultChan <- apiResult{resp: resp, err: err}
	}()

	// Wait for result or timeout
	select {
	case result := <-resultChan:
		totalAPITime := time.Since(apiStartTime)

		if result.err != nil {
			logger.Error("Gemini API request failed",
				"error", result.err,
				"duration", totalAPITime)
			return nil, fmt.Errorf("failed to generate AI response: %v", result.err)
		}

		logger.Info("Received response from Gemini API", "duration", totalAPITime)
		return ai.processAPIResponse(result.resp)

	case <-ctx.Done():
		totalTime := time.Since(apiStartTime)
		logger.Error("Gemini API request timed out",
			"timeout", timeout,
			"actual_duration", totalTime)
		return nil, fmt.Errorf("API request timed out after %v", timeout)
	}
}
func (ai *AIMapper) processAPIResponse(resp *genai.GenerateContentResponse) ([]AIMapping, error) {
	logger.Info("Processing AI response")

	if len(resp.Candidates) == 0 {
		logger.Error("No response candidates received from Gemini API")
		return nil, fmt.Errorf("no response generated from AI")
	}

	if len(resp.Candidates[0].Content.Parts) == 0 {
		logger.Error("No content parts in AI response")
		return nil, fmt.Errorf("no response generated from AI")
	}

	// Extract text from response
	var responseText string
	for i, part := range resp.Candidates[0].Content.Parts {
		if textPart, ok := part.(genai.Text); ok {
			partText := string(textPart)
			responseText += partText
			logger.Info("API response part", "part_index", i, "content", partText)
		} else {
			logger.Warn("Non-text part in response", "index", i, "type", fmt.Sprintf("%T", part))
		}
	}

	// Log the COMPLETE raw response
	logger.Info("COMPLETE RAW API RESPONSE", "full_response", responseText)

	// Parse the response
	mappings, err := ai.parseMappingResponse(responseText)
	if err != nil {
		logger.Error("Failed to parse AI response", "error", err)
		return nil, fmt.Errorf("failed to parse AI response: %v", err)
	}

	logger.Info("AI response parsed", "total_mappings_found", len(mappings))
	return mappings, nil
}

// buildMappingPrompt creates a prompt for the AI to generate column mappings
func (ai *AIMapper) buildMappingPrompt(scannedColumns, targetColumns []string) string {
	logger.Debug("Building AI prompt", "scanned_count", len(scannedColumns), "target_count", len(targetColumns))

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

	logger.Debug("Prompt built successfully", "final_length", len(prompt))
	return prompt
}

// parseMappingResponse parses the AI response into structured mappings
func (ai *AIMapper) parseMappingResponse(response string) ([]AIMapping, error) {
	logger.Info("Parsing AI response", "response_length", len(response))

	var mappings []AIMapping
	lines := strings.Split(strings.TrimSpace(response), "\n")

	logger.Info("Processing response lines", "total_lines", len(lines))

	for lineNum, line := range lines {
		line = strings.TrimSpace(line)
		logger.Info("Processing line", "line_num", lineNum+1, "content", line)

		if line == "" || strings.HasPrefix(line, "ScannedColumn|") {
			logger.Info("Skipping header/empty line", "line_num", lineNum+1)
			continue
		}

		parts := strings.Split(line, "|")
		if len(parts) != 3 {
			logger.Info("Skipping invalid format line", "line_num", lineNum+1, "parts_count", len(parts), "content", line)
			continue
		}

		scannedCol := strings.TrimSpace(parts[0])
		targetCol := strings.TrimSpace(parts[1])
		confidenceStr := strings.TrimSpace(parts[2])

		// Parse confidence
		var confidence float64
		if _, err := fmt.Sscanf(confidenceStr, "%f", &confidence); err != nil {
			logger.Info("Failed to parse confidence", "line_num", lineNum+1, "confidence_str", confidenceStr)
			confidence = 0.0
		}

		logger.Info("Parsed mapping",
			"line_num", lineNum+1,
			"scanned", scannedCol,
			"target", targetCol,
			"confidence", confidence)

		// Log what we're filtering out
		if targetCol == "NO_MATCH" {
			logger.Info("Filtering out NO_MATCH", "scanned", scannedCol)
			continue
		}

		if confidence < 0.6 {
			logger.Info("Filtering out low confidence", "scanned", scannedCol, "confidence", confidence)
			continue
		}

		mappings = append(mappings, AIMapping{
			ScannedColumn: scannedCol,
			TargetColumn:  targetCol,
			Confidence:    confidence,
		})
		logger.Info("Added mapping", "scanned", scannedCol, "target", targetCol, "confidence", confidence)
	}

	logger.Info("Parsing completed", "final_mappings", len(mappings))
	return mappings, nil
}

// GetGeminiAPIKey gets the API key from environment variable
func GetGeminiAPIKey() string {
	apiKey := os.Getenv("GEMINI_API_KEY")
	if apiKey == "" {
		logger.Warn("GEMINI_API_KEY environment variable not set")
	} else {
		logger.Debug("GEMINI_API_KEY found", "length", len(apiKey))
	}
	return apiKey
}
