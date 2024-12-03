package main

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
)

func main() {
	inputDir := "./input"
	outputDir := "./output"

	// Ensure the output directory exists
	if err := os.MkdirAll(outputDir, os.ModePerm); err != nil {
		fmt.Printf("Error creating output directory: %v\n", err)
		return
	}

	// Walk through the input directory and process `.pptx` files
	err := filepath.Walk(inputDir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}

		// Skip directories
		if info.IsDir() {
			return nil
		}

		// Only process `.pptx` files
		ext := strings.ToLower(filepath.Ext(info.Name()))
		if ext != ".pptx" {
			return nil
		}

		outputFilePath := filepath.Join(outputDir, strings.TrimSuffix(info.Name(), ext)+".txt")
		if err := processPptx(path, outputFilePath); err != nil {
			fmt.Printf("Failed to process file %s: %v\n", path, err)
		}

		return nil
	})

	if err != nil {
		fmt.Printf("Error walking input directory: %v\n", err)
	}
	fmt.Println("Processing complete.")
}

// processPptx reads a `.pptx` file and writes its content to a `.txt` file
func processPptx(inputPath, outputPath string) error {
	zipReader, err := zip.OpenReader(inputPath)
	if err != nil {
		return fmt.Errorf("error opening pptx file: %v", err)
	}
	defer zipReader.Close()

	var content strings.Builder
	for _, file := range zipReader.File {
		// Only process slide files (located in "ppt/slides/")
		if strings.HasPrefix(file.Name, "ppt/slides/slide") && strings.HasSuffix(file.Name, ".xml") {
			rc, err := file.Open()
			if err != nil {
				return fmt.Errorf("error reading slide %s: %v", file.Name, err)
			}

			data, err := io.ReadAll(rc)
			rc.Close()
			if err != nil {
				return fmt.Errorf("error reading slide content: %v", err)
			}

			content.WriteString(extractTextFromXML(data) + "\n")
		}
	}

	// Write the extracted content to the output file
	return os.WriteFile(outputPath, []byte(content.String()), os.ModePerm)
}

// extractTextFromXML extracts text content from XML data
func extractTextFromXML(data []byte) string {
	content := string(data)
	var text strings.Builder

	inTag := false
	for _, r := range content {
		if r == '<' {
			inTag = true
		} else if r == '>' {
			inTag = false
		} else if !inTag {
			text.WriteRune(r)
		}
	}

	return strings.TrimSpace(text.String())
}
