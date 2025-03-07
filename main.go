package main

import (
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

const OutputPath = "docxfreed"

func main() {
	var cfg Args
	parseArgs(&cfg)
	validateArgs(cfg)

	if err := createOutputPath(cfg); err != nil {
		fmt.Println(err)
		os.Exit(1)
	}

	if cfg.inputFile != "" {
		op := filepath.Join(OutputPath, "unprotected-"+cfg.inputFile)
		if err := processDocx(cfg.inputFile, op); err != nil {
			fmt.Println(err)
			os.Exit(1)
		}
	} else {
		if err := batchProcessDocx(cfg); err != nil {
			fmt.Println(err)
			os.Exit(1)
		}
	}
}

type Args struct {
	inputFile string
	inputPath string
	depth     int
}

func parseArgs(args *Args) {
	flag.StringVar(&args.inputFile, "f", "", "protected .docx filename")
	flag.StringVar(&args.inputPath, "p", "", "directory of .docx files for batch operation")
	flag.IntVar(&args.depth, "d", 1, "recursion depth")
	flag.Parse()
}

func validateArgs(cfg Args) {
	if cfg.inputFile == "" && cfg.inputPath == "" {
		printUsage()
		os.Exit(0)
	}

	if cfg.inputFile != "" && !strings.HasSuffix(cfg.inputFile, ".docx") {
		fmt.Println("file must be a '.docx'")
		os.Exit(1)
	}

	if cfg.depth < 0 {
		fmt.Println("depth must be greater than 0")
		os.Exit(1)
	}
}

func printUsage() {
	fmt.Println(`
	Arguments:
		-f: Specify a single filename.
		-p: Specify a path to a directory of files.
		-d: *Optional - Depth of recursion when specifying a path. Default: 1

    Usage:
		Single File:
		docxfree -f [filename]

		Batch Operations:
		docxfree -p [path/to/files]
		docxfree -p [path/to/files] -d [recursion-depth]

    Examples:
		docxfree -f document.docx
		docxfree -p protected_docs/
		docxfree -p /path/to/file -d 3`)
}

func createOutputPath(cfg Args) error {
	var bd string
	if cfg.inputPath != "" {
		bd = cfg.inputPath
	} else {
		bd = "."
	}
	if err := os.MkdirAll(filepath.Join(bd, OutputPath), 0700); err != nil {
		return fmt.Errorf("creating output directory: %w", err)
	}
	return nil
}
