#!/usr/bin/env bash
# menu.sh - usage: source menu.sh

# =============================================================================
# STARK096 — ALDEA: Generador de Ficheros Excel para Constructora
# =============================================================================

# Colors for beautiful output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
PURPLE='\033[0;35m'
CYAN='\033[0;36m'
WHITE='\033[1;37m'
NC='\033[0m' # No Color

# Emojis
DOCKER="🐳"
BUILD="🏗️"
ENTER="🖥️"
STOP="🛑"
PLAY="▶️"
EXCEL="📊"
TOOLS="🔧"
EXIT="❌"
HOUSE="🏘️"
ROCKET="🚀"
CHECK="✅"
WARNING="⚠️"

# Project configuration
IMAGE_NAME="stark096-aldea"
CONTAINER_NAME="stark096_aldea_dev"
WORKDIR_HOST="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
CODE_DIR="$WORKDIR_HOST/excelGenerator"

is_sourced() { [[ "${BASH_SOURCE[0]}" != "${0}" ]]; }

# =============================================================================
# UTILITIES
# =============================================================================

show_header() {
    clear
    echo -e "${PURPLE}================================================================${NC}"
    echo -e "${PURPLE}         ${WHITE}STARK096 — ALDEA: Excel Generator${PURPLE}                 ${NC}"
    echo -e "${PURPLE}            ${CYAN}Generador de Ficheros para Constructora${PURPLE}          ${NC}"
    echo -e "${PURPLE}================================================================${NC}"
    echo ""
    echo -e "${YELLOW}Project Path:${NC} ${WORKDIR_HOST}"
    echo -e "${YELLOW}Branch:${NC} $(git -C "$WORKDIR_HOST" branch --show-current 2>/dev/null || echo 'N/A')"
    echo ""
}

pause() { read -rp "Press Enter to continue..." _; }

confirm() {
    local msg="${1:-Are you sure?}"
    read -rn1 -p "$msg [y/N]: " ans
    echo
    [[ "$ans" =~ ^[Yy]$ ]]
}

safe_run_python() {
    local script="$1"
    shift
    local args=("$@")
    # Ensure container is running before executing inside it
    if ! docker ps --format '{{.Names}}' | grep -Eq "^${CONTAINER_NAME}$"; then
        echo -e "${YELLOW}${WARNING} Container not running. Starting...${NC}"
        echo ""
        if ! start_container; then
            return 1
        fi
        echo ""
    fi
    # Run script inside the container (-i keeps stdin open for user input prompts)
    MSYS_NO_PATHCONV=1 docker exec -i "$CONTAINER_NAME" python3 -u "/app/$script" "${args[@]}"
}

# =============================================================================
# DOCKER ACTIONS
# =============================================================================

build_image() {
    # Stop and remove existing container
    if docker ps -a --format '{{.Names}}' | grep -Eq "^${CONTAINER_NAME}$"; then
        echo -e "${GREEN}🛑 Removing existing container ${CONTAINER_NAME}...${NC}"
        docker rm -f "$CONTAINER_NAME" >/dev/null 2>&1 || true
        echo -e "${GREEN}${CHECK} Container removed.${NC}"
        echo ""
    fi

    # Remove existing image
    if [ -n "$(docker images -q "$IMAGE_NAME:latest")" ]; then
        echo -e "${GREEN}🗑️  Removing existing image ${IMAGE_NAME}:latest...${NC}"
        docker rmi "$IMAGE_NAME:latest" >/dev/null 2>&1 || true
        echo -e "${GREEN}${CHECK} Image removed.${NC}"
        echo ""
    fi

    # Build fresh image
    echo -e "${GREEN}${BUILD} Building image ${IMAGE_NAME}:latest...${NC}"
    echo ""
    if ! docker build -t "$IMAGE_NAME:latest" "$WORKDIR_HOST"; then
        echo ""
        echo -e "${RED}${EXIT} Build failed.${NC}"
        pause
        return 1
    fi
    echo ""
    echo -e "${GREEN}${CHECK} Image built.${NC}"
    echo ""

    # Start fresh container
    echo -e "${GREEN}${ROCKET} Starting fresh container...${NC}"
    start_container
    echo ""
    echo -e "${GREEN}${CHECK} Done — image built and container running.${NC}"
    pause
}

start_container() {
    if [ -z "$(docker images -q "$IMAGE_NAME:latest")" ]; then
        echo -e "${YELLOW}${WARNING} Image $IMAGE_NAME:latest not found. Building image now...${NC}"
        echo ""
        docker build -t "$IMAGE_NAME:latest" "$WORKDIR_HOST"
    fi

    if [ -z "$(docker images -q "$IMAGE_NAME:latest")" ]; then
        echo -e "${RED}${EXIT} Unable to build image. Cannot start container.${NC}"
        return 1
    fi

    if docker ps -a --format '{{.Names}}' | grep -Eq "^${CONTAINER_NAME}$"; then
        local running
        running=$(docker inspect -f '{{.State.Running}}' "$CONTAINER_NAME" 2>/dev/null || echo false)
        if [ "$running" = "true" ]; then
            echo -e "${GREEN}${CHECK} Container is already running.${NC}"
        else
            docker start "$CONTAINER_NAME" >/dev/null 2>&1 || true
            echo -e "${GREEN}${CHECK} Container started.${NC}"
        fi
    else
        local docker_args=(--name "$CONTAINER_NAME" -d -v "$CODE_DIR":/app)
        docker_args+=("$IMAGE_NAME:latest" bash -lc "tail -f /dev/null")
        if MSYS_NO_PATHCONV=1 docker run "${docker_args[@]}" >/dev/null 2>&1; then
            echo -e "${GREEN}${CHECK} Container created and started.${NC}"
        else
            echo -e "${RED}${EXIT} Failed to create container.${NC}"
            return 1
        fi
    fi
}

enter_container() {
    echo -e "${GREEN}${ENTER} Connecting to container shell...${NC}"
    echo ""
    if ! start_container; then
        pause
        return 1
    fi
    if docker exec -it "$CONTAINER_NAME" bash; then
        return 0
    fi
    echo -e "${YELLOW}${WARNING} Interactive shell failed; trying non-TTY fallback...${NC}"
    docker exec -i "$CONTAINER_NAME" bash
}

stop_remove_container() {
    if docker ps -a --format '{{.Names}}' | grep -Eq "^${CONTAINER_NAME}$"; then
        echo -e "${GREEN}🛑 Stopping and removing container ${CONTAINER_NAME}...${NC}"
        docker rm -f "$CONTAINER_NAME" >/dev/null 2>&1 || true
        echo -e "${GREEN}${CHECK} Container stopped and removed.${NC}"
    else
        echo -e "${YELLOW}${WARNING} No container named ${CONTAINER_NAME} found.${NC}"
    fi
    pause
}

show_container_status() {
    echo -e "${CYAN}📋 Container Status:${NC}"
    echo ""
    echo -e "${WHITE}Named container (${CONTAINER_NAME}):${NC}"
    docker ps -a --filter "name=${CONTAINER_NAME}" --format "table {{.Names}}\t{{.Status}}\t{{.Image}}" 2>/dev/null || echo "  (none)"
    echo ""
    echo -e "${WHITE}Image (${IMAGE_NAME}):${NC}"
    docker images "$IMAGE_NAME" --format "table {{.Repository}}\t{{.Tag}}\t{{.Size}}\t{{.CreatedSince}}" 2>/dev/null || echo "  (none)"
    echo ""
    pause
}

# =============================================================================
# SUBMENUS
# =============================================================================

show_container_menu() {
    while true; do
        show_header
        echo -e "${CYAN}================================================================${NC}"
        echo -e "${CYAN}                  ${WHITE}CONTAINER MANAGEMENT${CYAN}                      ${NC}"
        echo -e "${CYAN}              ${WHITE}Docker Image & Container Control${CYAN}               ${NC}"
        echo -e "${CYAN}================================================================${NC}"
        echo ""
        echo -e "${WHITE}${DOCKER} Docker Actions:${NC}"
        echo ""
        echo -e "${CYAN} ${GREEN}1.${NC} ${BUILD} Build Image (full rebuild)                        ${NC}"
        echo -e "      ↳ Remove container+image, build fresh, start          ${NC}"
        echo ""
        echo -e "${CYAN} ${GREEN}2.${NC} ${ROCKET} Start Container                                   ${NC}"
        echo -e "      ↳ Create and start ${CONTAINER_NAME}                  ${NC}"
        echo ""
        echo -e "${CYAN} ${GREEN}3.${NC} ${ENTER} Enter Container Shell                            ${NC}"
        echo -e "      ↳ Exec into running ${CONTAINER_NAME}                 ${NC}"
        echo ""
        echo -e "${CYAN} ${GREEN}4.${NC} 🛑 Stop / Remove Container                           ${NC}"
        echo -e "      ↳ Force-remove the named container                  ${NC}"
        echo ""
        echo -e "${CYAN} ${GREEN}5.${NC} 📋 Container Status                                  ${NC}"
        echo -e "      ↳ Show image and container info                     ${NC}"
        echo ""
        echo -e "${CYAN} ${GREEN}0.${NC} ${EXIT} Back to Main Menu                               ${NC}"
        echo -e "${CYAN}================================================================${NC}"
        echo ""
        read -rp "" opc
        case "$opc" in
            1) show_header; build_image ;;
            2) show_header; start_container; pause ;;
            3) enter_container; pause ;;
            4) show_header; stop_remove_container ;;
            5) show_header; show_container_status ;;
            0) return ;;
            *) echo -e "${RED}${EXIT} Invalid option. Please choose 0-5.${NC}"; pause ;;
        esac
    done
}

show_excel_menu() {
    while true; do
        show_header
        echo -e "${CYAN}================================================================${NC}"
        echo -e "${CYAN}                      ${WHITE}EXCEL GENERATOR${CYAN}                      ${NC}"
        echo -e "${CYAN}           ${WHITE}Generación de Ficheros Parametrizados${CYAN}            ${NC}"
        echo -e "${CYAN}================================================================${NC}"
        echo ""
        echo -e "${WHITE}${EXCEL} Available Options:${NC}"
        echo ""
        echo -e "${CYAN} ${GREEN}1.${NC} ${HOUSE} Generate Excel Template                          ${NC}"
        echo -e "      ↳ Crear plantilla base para constructora            ${NC}"
        echo ""
        echo -e "${CYAN} ${GREEN}0.${NC} ${EXIT} Back to Main Menu                               ${NC}"
        echo -e "${CYAN}================================================================${NC}"
        echo ""
        read -rp "" opc
        case "$opc" in
            1)
                show_header
                echo -e "${GREEN}${EXCEL} Generating Excel template...${NC}"
                echo ""
                safe_run_python excel_template_generator.py
                exit_code=$?
                echo ""
                if [ $exit_code -eq 0 ]; then
                    echo -e "${GREEN}${CHECK} Template generated successfully!${NC}"
                else
                    echo -e "${RED}${EXIT} Template generation failed (exit code: $exit_code)${NC}"
                fi
                pause
                ;;
            0) return ;;
            *) echo -e "${RED}${EXIT} Invalid option. Please choose 0-1.${NC}"; pause ;;
        esac
    done
}

# =============================================================================
# MAIN MENU
# =============================================================================

show_main_menu() {
    echo -e "${CYAN}================================================================${NC}"
    echo -e "${CYAN}                        ${WHITE}MAIN MENU${CYAN}                            ${NC}"
    echo -e "${CYAN}================================================================${NC}"
    echo ""
    echo -e "${WHITE}Select an option:${NC}"
    echo ""
    echo -e "${CYAN} ${GREEN}1.${NC} ${TOOLS} Container Management                               ${NC}"
    echo -e "      ↳ Build, enter, stop, and inspect containers        ${NC}"
    echo ""
    echo -e "${CYAN} ${GREEN}2.${NC} ${EXCEL} Excel Generator                                    ${NC}"
    echo -e "      ↳ Generate parametric Excel files for construction  ${NC}"
    echo ""
    echo -e "${CYAN} ${GREEN}0.${NC} ${EXIT} Exit                                               ${NC}"
    echo -e "${CYAN}================================================================${NC}"
    echo ""
    read -rp "" opc
    case "$opc" in
        1) show_container_menu ;;
        2) show_excel_menu ;;
        0) echo -e "${GREEN}${EXIT} Goodbye!${NC}"; return 0 ;;
        *) echo -e "${RED}${EXIT} Invalid option. Please choose 0-2.${NC}"; pause ;;
    esac
}

# =============================================================================
# ENTRY POINT
# =============================================================================

if ! is_sourced; then
    echo -e "${RED}❌ Error: This script must be sourced, not executed.${NC}"
    echo -e "${YELLOW}Usage: ${WHITE}source menu.sh${NC}"
    exit 1
fi

# Main loop
while true; do
    show_header
    show_main_menu
    [ $? -eq 0 ] && break
done
