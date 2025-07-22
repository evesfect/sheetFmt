package mapping

import (
	"fmt"
	"math"
	"strings"

	tea "github.com/charmbracelet/bubbletea"
	"github.com/charmbracelet/lipgloss"
)

// UI States
type state int

const (
	stateSelectScanned state = iota
	stateSelectTarget
	stateConfirm
)

// UIConfig represents UI configuration settings
type UIConfig struct {
	ColumnsPerRow int
	RowsPerPage   int
}

// Model represents the TUI model
type model struct {
	scannedColumns []string
	targetColumns  []string
	mappings       map[string]string // scanned -> target
	ignored        map[string]bool   // scanned -> ignored

	// UI state
	state          state
	currentScanned string

	// Grid navigation for scanned columns
	page         int
	row          int
	col          int
	colsPerRow   int
	rowsPerPage  int
	itemsPerPage int

	// Target selection
	targetCursor  int
	targetPage    int
	targetPerPage int

	// Screen dimensions
	width  int
	height int

	// Progress tracking
	mapped int
	total  int

	// Styling
	titleStyle    lipgloss.Style
	selectedStyle lipgloss.Style
	normalStyle   lipgloss.Style
	helpStyle     lipgloss.Style
	progressStyle lipgloss.Style
	mappedStyle   lipgloss.Style
	ignoredStyle  lipgloss.Style
}

// Initialize the model with config
func initialModel(scannedColumns, targetColumns []string, uiConfig UIConfig) model {
	return model{
		scannedColumns: scannedColumns,
		targetColumns:  targetColumns,
		mappings:       make(map[string]string),
		ignored:        make(map[string]bool),
		state:          stateSelectScanned,
		page:           0,
		row:            0,
		col:            0,
		colsPerRow:     uiConfig.ColumnsPerRow,
		rowsPerPage:    uiConfig.RowsPerPage,
		itemsPerPage:   uiConfig.ColumnsPerRow * uiConfig.RowsPerPage,
		targetCursor:   0,
		targetPage:     0,
		targetPerPage:  15,
		total:          len(scannedColumns),

		// ... styling stays the same
		titleStyle: lipgloss.NewStyle().
			Bold(true).
			Foreground(lipgloss.Color("205")).
			Align(lipgloss.Center),
		selectedStyle: lipgloss.NewStyle().
			Bold(true).
			Foreground(lipgloss.Color("170")).
			Background(lipgloss.Color("235")).
			Padding(0, 1),
		normalStyle: lipgloss.NewStyle().
			Foreground(lipgloss.Color("252")).
			Padding(0, 1),
		helpStyle: lipgloss.NewStyle().
			Foreground(lipgloss.Color("241")),
		progressStyle: lipgloss.NewStyle().
			Foreground(lipgloss.Color("205")).
			Bold(true),
		mappedStyle: lipgloss.NewStyle().
			Foreground(lipgloss.Color("40")).
			Padding(0, 1),
		ignoredStyle: lipgloss.NewStyle().
			Foreground(lipgloss.Color("240")).
			Strikethrough(true).
			Padding(0, 1),
	}
}

func (m model) Init() tea.Cmd {
	return nil
}

func (m model) Update(msg tea.Msg) (tea.Model, tea.Cmd) {
	switch msg := msg.(type) {
	case tea.WindowSizeMsg:
		m.width = msg.Width
		m.height = msg.Height

		// Only adjust target items per page based on height
		m.targetPerPage = m.height - 6
		if m.targetPerPage < 5 {
			m.targetPerPage = 5
		}
	case tea.KeyMsg:
		switch m.state {
		case stateSelectScanned:
			return m.updateSelectScanned(msg)
		case stateSelectTarget:
			return m.updateSelectTarget(msg)
		case stateConfirm:
			return m.updateConfirm(msg)
		}
	}
	return m, nil
}

func (m model) updateSelectScanned(msg tea.KeyMsg) (tea.Model, tea.Cmd) {
	switch msg.String() {
	case "ctrl+c", "q":
		return m, tea.Quit

	case "up", "k":
		if m.row > 0 {
			m.row--
		}

	case "down", "j":
		maxRow := m.getMaxRowForCurrentPage()
		if m.row < maxRow {
			m.row++
		}

	case "left", "h":
		if m.col > 0 {
			m.col--
		} else if m.page > 0 {
			// Go to previous page, rightmost column
			m.page--
			m.col = m.colsPerRow - 1
			// Adjust if the new position is out of bounds
			m.adjustPosition()
		}

	case "right", "l":
		maxCol := m.getMaxColForCurrentRow()
		if m.col < maxCol {
			m.col++
		} else if m.hasNextPage() {
			// Go to next page, leftmost column
			m.page++
			m.col = 0
			m.row = 0
		}

	case "enter":
		currentIdx := m.getCurrentIndex()
		if currentIdx < len(m.scannedColumns) {
			m.currentScanned = m.scannedColumns[currentIdx]
			m.state = stateSelectTarget
			m.targetCursor = 0
			m.targetPage = 0
		}

	case "i":
		// Toggle ignore for current column
		currentIdx := m.getCurrentIndex()
		if currentIdx < len(m.scannedColumns) {
			scanned := m.scannedColumns[currentIdx]
			if m.ignored[scanned] {
				delete(m.ignored, scanned)
				delete(m.mappings, scanned)
				m.mapped--
			} else {
				m.ignored[scanned] = true
				delete(m.mappings, scanned)
				m.mapped++
			}
		}

	case "n":
		// Move to next unmapped column
		m.moveToNextUnmapped()

	case "s":
		// Save and exit
		m.state = stateConfirm
	}
	return m, nil
}

func (m model) updateSelectTarget(msg tea.KeyMsg) (tea.Model, tea.Cmd) {
	switch msg.String() {
	case "ctrl+c", "q":
		return m, tea.Quit
	case "esc":
		m.state = stateSelectScanned
	case "up", "k":
		if m.targetCursor > 0 {
			m.targetCursor--
		} else if m.targetPage > 0 {
			m.targetPage--
			m.targetCursor = m.targetPerPage - 1
		}
	case "down", "j":
		maxCursor := m.getMaxTargetCursor()
		if m.targetCursor < maxCursor {
			m.targetCursor++
		} else if m.hasNextTargetPage() {
			m.targetPage++
			m.targetCursor = 0
		}
	case "left", "h":
		if m.targetPage > 0 {
			m.targetPage--
		}
	case "right", "l":
		if m.hasNextTargetPage() {
			m.targetPage++
		}
	case "enter":
		// Map the columns
		targetIdx := m.targetPage*m.targetPerPage + m.targetCursor
		if targetIdx < len(m.targetColumns) {
			target := m.targetColumns[targetIdx]

			// Remove any previous mapping for this scanned column
			if _, exists := m.mappings[m.currentScanned]; !exists {
				m.mapped++
			}

			m.mappings[m.currentScanned] = target
			delete(m.ignored, m.currentScanned)

			m.state = stateSelectScanned

			// Move to next unmapped column
			m.moveToNextUnmapped()
		}
	}
	return m, nil
}

func (m model) updateConfirm(msg tea.KeyMsg) (tea.Model, tea.Cmd) {
	switch msg.String() {
	case "ctrl+c", "q", "n":
		return m, tea.Quit
	case "y":
		return m, tea.Quit
	case "esc":
		m.state = stateSelectScanned
	}
	return m, nil
}

// Helper functions
func (m model) getCurrentIndex() int {
	return m.page*m.itemsPerPage + m.row*m.colsPerRow + m.col
}

func (m model) getMaxRowForCurrentPage() int {
	startOfPage := m.page * m.itemsPerPage
	remainingItems := len(m.scannedColumns) - startOfPage
	if remainingItems <= 0 {
		return 0
	}

	maxRowsNeeded := int(math.Ceil(float64(remainingItems) / float64(m.colsPerRow)))
	if maxRowsNeeded > m.rowsPerPage {
		return m.rowsPerPage - 1
	}
	return maxRowsNeeded - 1
}

func (m model) getMaxColForCurrentRow() int {
	startOfRow := m.page*m.itemsPerPage + m.row*m.colsPerRow
	endOfRow := startOfRow + m.colsPerRow
	if endOfRow > len(m.scannedColumns) {
		endOfRow = len(m.scannedColumns)
	}
	return (endOfRow - startOfRow) - 1
}

func (m model) hasNextPage() bool {
	return (m.page+1)*m.itemsPerPage < len(m.scannedColumns)
}

func (m model) hasNextTargetPage() bool {
	return (m.targetPage+1)*m.targetPerPage < len(m.targetColumns)
}

func (m model) getMaxTargetCursor() int {
	itemsOnPage := len(m.targetColumns) - m.targetPage*m.targetPerPage
	if itemsOnPage > m.targetPerPage {
		return m.targetPerPage - 1
	}
	return itemsOnPage - 1
}

func (m *model) adjustPosition() {
	// Ensure current position is valid
	currentIdx := m.getCurrentIndex()
	if currentIdx >= len(m.scannedColumns) {
		// Move to last valid position
		m.moveToLastValidPosition()
	}
}

func (m *model) moveToLastValidPosition() {
	if len(m.scannedColumns) == 0 {
		return
	}
	lastIdx := len(m.scannedColumns) - 1
	m.page = lastIdx / m.itemsPerPage
	remainder := lastIdx % m.itemsPerPage
	m.row = remainder / m.colsPerRow
	m.col = remainder % m.colsPerRow
}

func (m *model) moveToNextUnmapped() {
	// Safety check - prevent division by zero
	if m.itemsPerPage == 0 || m.colsPerRow == 0 {
		return
	}

	currentIdx := m.getCurrentIndex()

	// First search from current position forward
	for i := currentIdx + 1; i < len(m.scannedColumns); i++ {
		scanned := m.scannedColumns[i]
		if _, mapped := m.mappings[scanned]; !mapped && !m.ignored[scanned] {
			m.page = i / m.itemsPerPage
			remainder := i % m.itemsPerPage
			m.row = remainder / m.colsPerRow
			m.col = remainder % m.colsPerRow
			return
		}
	}

	// If no unmapped found after cursor, search from beginning
	for i := 0; i < currentIdx; i++ {
		scanned := m.scannedColumns[i]
		if _, mapped := m.mappings[scanned]; !mapped && !m.ignored[scanned] {
			m.page = i / m.itemsPerPage
			remainder := i % m.itemsPerPage
			m.row = remainder / m.colsPerRow
			m.col = remainder % m.colsPerRow
			return
		}
	}

	// If no unmapped columns found anywhere, stay at current position
	// (This means all columns are either mapped or ignored)
}

func (m model) View() string {
	switch m.state {
	case stateSelectScanned:
		return m.viewSelectScanned()
	case stateSelectTarget:
		return m.viewSelectTarget()
	case stateConfirm:
		return m.viewConfirm()
	}
	return ""
}

func (m model) viewSelectScanned() string {
	var b strings.Builder

	// Title
	title := m.titleStyle.Width(m.width).Render("Column Mapping Tool")
	b.WriteString(title)
	b.WriteString("\n\n")

	// Progress
	progress := fmt.Sprintf("Progress: %d/%d mapped (%d ignored)",
		len(m.mappings), m.total, len(m.ignored))
	b.WriteString(m.progressStyle.Render(progress))
	b.WriteString("\n\n")

	// Page info
	totalPages := int(math.Ceil(float64(len(m.scannedColumns)) / float64(m.itemsPerPage)))
	if totalPages == 0 {
		totalPages = 1
	}
	pageInfo := fmt.Sprintf("Page %d/%d", m.page+1, totalPages)
	b.WriteString(m.helpStyle.Render(pageInfo))
	b.WriteString("\n\n")

	// Calculate column width dynamically
	columnWidth := (m.width - 4) / m.colsPerRow // Account for padding and spacing
	if columnWidth < 10 {
		columnWidth = 10 // Minimum width
	}

	// Column grid - use configurable rows
	for row := 0; row < m.rowsPerPage; row++ {
		var rowItems []string
		for col := 0; col < m.colsPerRow; col++ {
			idx := m.page*m.itemsPerPage + row*m.colsPerRow + col
			if idx >= len(m.scannedColumns) {
				break
			}

			column := m.scannedColumns[idx]
			var style lipgloss.Style
			var displayText string

			// Create display text with mapping info
			if target, mapped := m.mappings[column]; mapped {
				displayText = fmt.Sprintf("%s → %s", column, target)
				style = m.mappedStyle
			} else if m.ignored[column] {
				displayText = fmt.Sprintf("%s (ignored)", column)
				style = m.ignoredStyle
			} else {
				displayText = column
				style = m.normalStyle
			}

			// Highlight if selected
			if row == m.row && col == m.col {
				style = m.selectedStyle
			}

			// Truncate based on calculated column width
			if len(displayText) > columnWidth-2 {
				displayText = displayText[:columnWidth-5] + "..."
			}

			// Use calculated width for consistent spacing
			displayText = fmt.Sprintf("%-*s", columnWidth-2, displayText)

			rowItems = append(rowItems, style.Render(displayText))
		}

		if len(rowItems) > 0 {
			b.WriteString(lipgloss.JoinHorizontal(lipgloss.Top, rowItems...))
			b.WriteString("\n")
		}
	}

	b.WriteString("\n")

	// Help
	help := "↑↓←→: navigate | Enter: select | i: ignore | n: next unmapped | s: save | q: quit"
	b.WriteString(m.helpStyle.Render(help))

	return b.String()
}

func (m model) viewSelectTarget() string {
	var b strings.Builder

	// Title
	title := fmt.Sprintf("Map '%s' to target column:", m.currentScanned)
	b.WriteString(m.titleStyle.Render(title))
	b.WriteString("\n\n")

	// Page info
	totalPages := int(math.Ceil(float64(len(m.targetColumns)) / float64(m.targetPerPage)))
	if totalPages == 0 {
		totalPages = 1
	}
	pageInfo := fmt.Sprintf("Page %d/%d", m.targetPage+1, totalPages)
	b.WriteString(m.helpStyle.Render(pageInfo))
	b.WriteString("\n\n")

	// Target columns list
	start := m.targetPage * m.targetPerPage
	end := start + m.targetPerPage
	if end > len(m.targetColumns) {
		end = len(m.targetColumns)
	}

	for i := start; i < end; i++ {
		column := m.targetColumns[i]
		localIndex := i - start

		if localIndex == m.targetCursor {
			b.WriteString(m.selectedStyle.Render("> " + column))
		} else {
			b.WriteString(m.normalStyle.Render("  " + column))
		}
		b.WriteString("\n")
	}

	b.WriteString("\n")

	// Help
	help := "↑↓: navigate | ←→: prev/next page | Enter: select | Esc: back | q: quit"
	b.WriteString(m.helpStyle.Render(help))

	return b.String()
}

func (m model) viewConfirm() string {
	var b strings.Builder

	b.WriteString(m.titleStyle.Render("Save Mapping Configuration?"))
	b.WriteString("\n\n")

	// Summary
	b.WriteString(fmt.Sprintf("Total columns: %d\n", m.total))
	b.WriteString(fmt.Sprintf("Mapped: %d\n", len(m.mappings)))
	b.WriteString(fmt.Sprintf("Ignored: %d\n", len(m.ignored)))
	b.WriteString(fmt.Sprintf("Unmapped: %d\n", m.total-len(m.mappings)-len(m.ignored)))
	b.WriteString("\n")

	b.WriteString(m.helpStyle.Render("y/n to confirm, Esc to go back"))

	return b.String()
}

// RunMappingTUI starts the interactive mapping interface
func RunMappingTUI(scannedColumnsFile, targetColumnsFile, outputMappingFile string, uiConfig UIConfig) error {
	scannedColumns, err := ReadColumnsFromFile(scannedColumnsFile)
	if err != nil {
		return fmt.Errorf("failed to read scanned columns: %v", err)
	}

	if len(scannedColumns) == 0 {
		return fmt.Errorf("no scanned columns found in %s", scannedColumnsFile)
	}

	// Read target columns
	targetColumns, err := ReadColumnsFromFile(targetColumnsFile)
	if err != nil {
		return fmt.Errorf("failed to read target columns: %v", err)
	}

	if len(targetColumns) == 0 {
		return fmt.Errorf("no target columns found in %s", targetColumnsFile)
	}

	// Initialize the TUI model with config
	m := initialModel(scannedColumns, targetColumns, uiConfig)

	// Load existing mappings if the file exists
	if existingConfig, err := LoadFromFile(outputMappingFile); err == nil {
		fmt.Printf("📂 Loading existing mappings from %s\n", outputMappingFile)

		// Apply existing mappings to the model
		for _, mapping := range existingConfig.Mappings {
			if mapping.IsIgnored {
				m.ignored[mapping.ScannedColumn] = true
				m.mapped++
			} else if mapping.TargetColumn != "" {
				m.mappings[mapping.ScannedColumn] = mapping.TargetColumn
				m.mapped++
			}
		}

		fmt.Printf("✓ Loaded %d existing mappings (%d mapped, %d ignored)\n",
			len(existingConfig.Mappings), len(m.mappings), len(m.ignored))

		// Move to first unmapped column
		m.moveToNextUnmapped()
	} else {
		fmt.Printf("📝 No existing mappings found, starting fresh\n")
	}

	// Run the TUI
	p := tea.NewProgram(m, tea.WithAltScreen())
	finalModel, err := p.Run()
	if err != nil {
		return fmt.Errorf("error running TUI: %v", err)
	}

	// Get the final model
	final := finalModel.(model)

	// Check if user wants to save (if they pressed 's' and confirmed with 'y')
	if final.state == stateConfirm {
		// Create mapping config
		config := &MappingConfig{}

		// Add all mappings
		for scanned, target := range final.mappings {
			config.Mappings = append(config.Mappings, ColumnMapping{
				ScannedColumn: scanned,
				TargetColumn:  target,
				IsIgnored:     false,
			})
		}

		// Add ignored columns
		for scanned := range final.ignored {
			config.Mappings = append(config.Mappings, ColumnMapping{
				ScannedColumn: scanned,
				TargetColumn:  "",
				IsIgnored:     true,
			})
		}

		// Save the mapping configuration
		err = config.SaveToFile(outputMappingFile)
		if err != nil {
			return fmt.Errorf("failed to save mapping configuration: %v", err)
		}

		fmt.Printf("✓ Mapping configuration saved to: %s\n", outputMappingFile)
		fmt.Printf("✓ Mapped %d columns, ignored %d columns\n", len(final.mappings), len(final.ignored))
	}

	return nil
}
