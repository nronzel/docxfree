package main

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
)

func findFileInZip(reader *zip.ReadCloser, path string) *zip.File {
	for _, file := range reader.File {
		if file.Name == path {
			return file
		}
	}

	return nil
}

func readZipFile(file *zip.File) ([]byte, error) {
	rc, err := file.Open()
	if err != nil {
		return nil, fmt.Errorf("opening file in archive: %w", err)
	}
	defer rc.Close()

	content, err := io.ReadAll(rc)
	if err != nil {
		return nil, fmt.Errorf("reading content: %w", err)
	}

	return content, nil
}

func processDocx(inputPath, outputPath string) error {
	reader, err := zip.OpenReader(inputPath)
	if err != nil {
		return fmt.Errorf("opening docx file: %w", err)
	}
	defer reader.Close()

	settingsPath := "word/settings.xml"

	settingsFile := findFileInZip(reader, settingsPath)
	if settingsFile == nil {
		fmt.Println("No settings.xml found in the document. No changes were made.")
		return nil
	}

	settingsContent, err := readZipFile(settingsFile)
	if err != nil {
		return fmt.Errorf("reading settings.xml: %w", err)
	}

	modifiedSettings, protectionFound, err := removeProtectionFromXML(settingsContent)
	if err != nil {
		return fmt.Errorf("processing settings.xml: %w", err)
	}

	if !protectionFound {
		fmt.Println("No document protection found. No changes were made.")
		return nil
	}

	err = createUnprotectedCopy(reader, outputPath, settingsPath, modifiedSettings)
	if err != nil {
		return fmt.Errorf("creating unprotected copy: %w", err)
	}

	fmt.Printf("Successfully removed protection and created: %s\n", outputPath)
	return nil
}

func removeProtectionFromXML(xmlData []byte) ([]byte, bool, error) {
	decoder := xml.NewDecoder(bytes.NewReader(xmlData))
	var buf bytes.Buffer
	encoder := xml.NewEncoder(&buf)
	protectionFound := false

	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, false, fmt.Errorf("decoding XML: %w", err)
		}

		if startElement, ok := token.(xml.StartElement); ok {
			if startElement.Name.Local == "documentProtection" {
				protectionFound = true
				if err := decoder.Skip(); err != nil {
					return nil, false, fmt.Errorf("skipping documentProtection: %w", err)
				}
				continue
			}
		}

		if err := encoder.EncodeToken(token); err != nil {
			return nil, false, fmt.Errorf("encoding XML: %w", err)
		}
	}

	if err := encoder.Flush(); err != nil {
		return nil, false, fmt.Errorf("flushing encoder: %w", err)
	}

	// if there is no protection, return the original data
	if !protectionFound {
		return xmlData, false, nil
	}

	return buf.Bytes(), true, nil
}

func batchProcessDocx(cfg Args) error {
	return processDirectory(cfg.inputPath, 1, cfg.depth)
}

func processDirectory(dirPath string, currentDepth, maxDepth int) error {
	files, err := os.ReadDir(dirPath)
	if err != nil {
		return fmt.Errorf("reading directory %s: %w", dirPath, err)
	}

	outputDir := filepath.Join(dirPath, OutputPath)
	if err := os.MkdirAll(outputDir, 0700); err != nil {
		return fmt.Errorf("creating output directory %s: %w", outputDir, err)
	}

	for _, file := range files {
		fullPath := filepath.Join(dirPath, file.Name())

		if file.IsDir() {
			if file.Name() == OutputPath {
				continue
			}

			if currentDepth < maxDepth {
				if err := processDirectory(fullPath, currentDepth+1, maxDepth); err != nil {
					return err
				}
			}
			continue
		}

		if strings.HasSuffix(file.Name(), ".docx") {
			outputPath := filepath.Join(outputDir, "unprotected-"+file.Name())
			fmt.Printf("Processing: %s\n", fullPath)

			if err := processDocx(fullPath, outputPath); err != nil {
				return err
			}
		}
	}
	return nil
}

func createUnprotectedCopy(
	reader *zip.ReadCloser,
	outputPath, settingsPath string,
	modifiedSettings []byte,
) error {
	outputFile, err := os.Create(outputPath)
	if err != nil {
		return fmt.Errorf("creating output file: %w", err)
	}
	defer outputFile.Close()

	zipWriter := zip.NewWriter(outputFile)
	defer zipWriter.Close()

	for _, file := range reader.File {
		writer, err := zipWriter.CreateHeader(&zip.FileHeader{
			Name:   file.Name,
			Method: file.Method,
		})
		if err != nil {
			return fmt.Errorf("creating file in output archive: %w", err)
		}

		if file.Name == settingsPath {
			if _, err := writer.Write(modifiedSettings); err != nil {
				return fmt.Errorf("writing modified settings.xml: %w", err)
			}
		} else {
			rc, err := file.Open()
			if err != nil {
				return fmt.Errorf("opening file in archive: %w", err)
			}

			if _, err := io.Copy(writer, rc); err != nil {
				rc.Close()
				return fmt.Errorf("copying file content for %s: %w", file.Name, err)
			}
			rc.Close()
		}
	}
	return nil
}
