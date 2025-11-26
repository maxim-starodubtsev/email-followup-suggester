#!/bin/bash

# Docker Test Runner Script
# This script provides easy commands to run tests in Docker

set -e

# Colors for output
GREEN='\033[0;32m'
BLUE='\033[0;34m'
RED='\033[0;31m'
NC='\033[0m' # No Color

echo -e "${BLUE}=== Followup Suggester Docker Test Runner ===${NC}\n"

# Function to print usage
usage() {
    echo "Usage: ./docker-test.sh [command]"
    echo ""
    echo "Commands:"
    echo "  build       - Build the Docker image"
    echo "  test        - Run all tests once"
    echo "  watch       - Run tests in watch mode"
    echo "  coverage    - Run tests with coverage report"
    echo "  shell       - Open a shell in the container"
    echo "  clean       - Remove Docker images and volumes"
    echo "  rebuild     - Clean and rebuild everything"
    echo ""
    exit 1
}

# Check if Docker is installed
if ! command -v docker &> /dev/null; then
    echo -e "${RED}Error: Docker is not installed${NC}"
    echo "Please install Docker Desktop from: https://www.docker.com/products/docker-desktop"
    exit 1
fi

# Check if Docker is running
if ! docker info &> /dev/null; then
    echo -e "${RED}Error: Docker is not running${NC}"
    echo "Please start Docker Desktop and try again"
    exit 1
fi

# Parse command
COMMAND=${1:-test}

case "$COMMAND" in
    build)
        echo -e "${BLUE}Building Docker image...${NC}"
        docker-compose build test
        echo -e "${GREEN}✓ Build complete${NC}"
        ;;
    
    test)
        echo -e "${BLUE}Running tests...${NC}"
        docker-compose run --rm test
        echo -e "${GREEN}✓ Tests complete${NC}"
        ;;
    
    watch)
        echo -e "${BLUE}Starting tests in watch mode...${NC}"
        echo -e "${BLUE}Press Ctrl+C to stop${NC}\n"
        docker-compose run --rm test-watch
        ;;
    
    coverage)
        echo -e "${BLUE}Running tests with coverage...${NC}"
        docker-compose run --rm test-coverage
        echo -e "${GREEN}✓ Coverage report generated in ./coverage${NC}"
        ;;
    
    shell)
        echo -e "${BLUE}Opening shell in container...${NC}"
        docker-compose run --rm test sh
        ;;
    
    clean)
        echo -e "${BLUE}Cleaning Docker resources...${NC}"
        docker-compose down -v
        docker rmi email-followup-suggester_test 2>/dev/null || true
        echo -e "${GREEN}✓ Cleanup complete${NC}"
        ;;
    
    rebuild)
        echo -e "${BLUE}Rebuilding from scratch...${NC}"
        docker-compose down -v
        docker-compose build --no-cache test
        docker-compose run --rm test
        echo -e "${GREEN}✓ Rebuild and test complete${NC}"
        ;;
    
    *)
        echo -e "${RED}Unknown command: $COMMAND${NC}\n"
        usage
        ;;
esac

