"use client";

import { useEffect, useMemo, useRef, useState, useSyncExternalStore, type PointerEvent as ReactPointerEvent, type ReactNode } from "react";

import type { ReportAnnotation, ReportAnnotationTool } from "@/lib/annotations/types";
import {
  ANNOTATION_BORDER_COLORS,
  ANNOTATION_FONT_SIZES,
  ANNOTATION_SURFACE_COLORS,
  ANNOTATION_TEXT_COLORS,
} from "@/lib/annotations/types";

type AnnotationSaveState = "idle" | "dirty" | "saving" | "saved" | "conflict" | "error";
type AnnotationMenuId = "fontSize" | "textColor" | "fillColor" | "borderColor";
type ToolbarPosition = { left: number; top: number };
type ToolbarDragState = {
  pointerId: number;
  originClientX: number;
  originClientY: number;
  startLeft: number;
  startTop: number;
  width: number;
  height: number;
};

const TOOLBAR_POSITION_STORAGE_KEY = "ta-it-reporting-annotation-toolbar-position";

function parseToolbarPosition(raw: string | null): ToolbarPosition | null {
  if (!raw) {
    return null;
  }

  try {
    const parsed = JSON.parse(raw) as Partial<ToolbarPosition>;
    if (typeof parsed.left !== "number" || typeof parsed.top !== "number") {
      return null;
    }

    return {
      left: parsed.left,
      top: parsed.top,
    };
  } catch {
    return null;
  }
}

function getStoredToolbarPositionSnapshot() {
  try {
    return window.localStorage.getItem(TOOLBAR_POSITION_STORAGE_KEY);
  } catch {
    return null;
  }
}

function subscribeToToolbarPosition(onStoreChange: () => void) {
  function handleStorage(event: StorageEvent) {
    if (event.storageArea === window.localStorage && (event.key === TOOLBAR_POSITION_STORAGE_KEY || event.key == null)) {
      onStoreChange();
    }
  }

  window.addEventListener("storage", handleStorage);
  return () => window.removeEventListener("storage", handleStorage);
}

interface ReportAnnotationToolbarProps {
  activeTool: ReportAnnotationTool;
  canEdit: boolean;
  saveMessage: string;
  saveState: AnnotationSaveState;
  selectedAnnotation: ReportAnnotation | null;
  onDeleteSelected: () => void;
  onToolChange: (tool: ReportAnnotationTool) => void;
  onUpdateSelected: (updater: (annotation: ReportAnnotation) => ReportAnnotation) => void;
}

function ToolIcon({ children }: { children: ReactNode }) {
  return <span className="annotation-tool-icon" aria-hidden="true">{children}</span>;
}

function ChevronIcon() {
  return (
    <svg fill="none" viewBox="0 0 12 12" xmlns="http://www.w3.org/2000/svg">
      <path d="M2.25 4.25 6 7.75l3.75-3.5" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" />
    </svg>
  );
}

function DragGripIcon() {
  return (
    <svg fill="none" viewBox="0 0 12 12" xmlns="http://www.w3.org/2000/svg">
      <circle cx="3" cy="3" r="1" fill="currentColor" />
      <circle cx="3" cy="6" r="1" fill="currentColor" />
      <circle cx="3" cy="9" r="1" fill="currentColor" />
      <circle cx="9" cy="3" r="1" fill="currentColor" />
      <circle cx="9" cy="6" r="1" fill="currentColor" />
      <circle cx="9" cy="9" r="1" fill="currentColor" />
    </svg>
  );
}

function MenuShell({
  isOpen,
  children,
  menu,
}: {
  isOpen: boolean;
  children: ReactNode;
  menu: ReactNode;
}) {
  return (
    <div className="annotation-menu-shell">
      {children}
      {isOpen ? <div className="annotation-menu">{menu}</div> : null}
    </div>
  );
}

export function ReportAnnotationToolbar({
  activeTool,
  canEdit,
  saveMessage,
  saveState,
  selectedAnnotation,
  onDeleteSelected,
  onToolChange,
  onUpdateSelected,
}: ReportAnnotationToolbarProps) {
  const rootRef = useRef<HTMLDivElement | null>(null);
  const [openMenu, setOpenMenu] = useState<AnnotationMenuId | null>(null);
  const persistedPositionSnapshot = useSyncExternalStore(subscribeToToolbarPosition, getStoredToolbarPositionSnapshot, () => null);
  const persistedPosition = useMemo(
    () => parseToolbarPosition(persistedPositionSnapshot),
    [persistedPositionSnapshot],
  );
  const [positionOverride, setPositionOverride] = useState<ToolbarPosition | null>(null);
  const position = positionOverride ?? persistedPosition;
  const [dragState, setDragState] = useState<ToolbarDragState | null>(null);

  useEffect(() => {
    function handlePointerDown(event: MouseEvent) {
      if (!rootRef.current?.contains(event.target as Node)) {
        setOpenMenu(null);
      }
    }

    document.addEventListener("mousedown", handlePointerDown);
    return () => document.removeEventListener("mousedown", handlePointerDown);
  }, []);

  useEffect(() => {
    if (!positionOverride) {
      return;
    }

    try {
      window.localStorage.setItem(TOOLBAR_POSITION_STORAGE_KEY, JSON.stringify(positionOverride));
    } catch {
      // localStorage is optional
    }
  }, [positionOverride]);

  useEffect(() => {
    function clampToViewport() {
      const root = rootRef.current;
      if (!root) {
        return;
      }

      setPositionOverride((current) => {
        const sourcePosition = current ?? persistedPosition;
        if (!sourcePosition) {
          return current;
        }

        const rect = root.getBoundingClientRect();
        const maxLeft = Math.max(12, window.innerWidth - rect.width - 12);
        const maxTop = Math.max(12, window.innerHeight - rect.height - 12);
        const nextLeft = Math.min(Math.max(12, sourcePosition.left), maxLeft);
        const nextTop = Math.min(Math.max(12, sourcePosition.top), maxTop);

        if (nextLeft === sourcePosition.left && nextTop === sourcePosition.top) {
          return current;
        }

        return {
          left: nextLeft,
          top: nextTop,
        };
      });
    }

    window.addEventListener("resize", clampToViewport);
    return () => window.removeEventListener("resize", clampToViewport);
  }, [persistedPosition]);

  useEffect(() => {
    if (!dragState) {
      return;
    }

    const currentDragState = dragState;

    function handlePointerMove(event: PointerEvent) {
      const deltaX = event.clientX - currentDragState.originClientX;
      const deltaY = event.clientY - currentDragState.originClientY;
      const maxLeft = Math.max(12, window.innerWidth - currentDragState.width - 12);
      const maxTop = Math.max(12, window.innerHeight - currentDragState.height - 12);

      setPositionOverride({
        left: Math.min(Math.max(12, currentDragState.startLeft + deltaX), maxLeft),
        top: Math.min(Math.max(12, currentDragState.startTop + deltaY), maxTop),
      });
    }

    function handlePointerUp(event: PointerEvent) {
      if (event.pointerId === currentDragState.pointerId) {
        setDragState(null);
      }
    }

    window.addEventListener("pointermove", handlePointerMove);
    window.addEventListener("pointerup", handlePointerUp);

    return () => {
      window.removeEventListener("pointermove", handlePointerMove);
      window.removeEventListener("pointerup", handlePointerUp);
    };
  }, [dragState]);

  if (!canEdit) {
    return null;
  }

  const toggleMenu = (menuId: AnnotationMenuId) => {
    setOpenMenu((current) => (current === menuId ? null : menuId));
  };

  const startDrag = (event: ReactPointerEvent<HTMLButtonElement>) => {
    const root = rootRef.current;
    if (!root) {
      return;
    }

    const rect = root.getBoundingClientRect();
    event.preventDefault();
    setOpenMenu(null);
    setDragState({
      pointerId: event.pointerId,
      originClientX: event.clientX,
      originClientY: event.clientY,
      startLeft: rect.left,
      startTop: rect.top,
      width: rect.width,
      height: rect.height,
    });
  };

  return (
    <div
      className="annotation-toolbar-shell"
      style={
        position
          ? {
              left: `${position.left}px`,
              top: `${position.top}px`,
              right: "auto",
            }
          : undefined
      }
    >
      <div className="annotation-toolbar annotation-toolbar-compact" ref={rootRef}>
        <div className="annotation-toolbar-group">
          <button
            aria-label="Drag toolbar"
            className="annotation-toolbar-grip"
            onPointerDown={startDrag}
            title="Drag toolbar"
            type="button"
          >
            <DragGripIcon />
          </button>
          <button
            aria-label="Select annotation"
            className="annotation-tool-button annotation-tool-button-icon"
            data-active={activeTool === "select"}
            onClick={() => onToolChange("select")}
            title="Select"
            type="button"
          >
            <ToolIcon>
              <svg fill="none" viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg">
                <path d="M4 3.5v12.75l3.9-3.3 2.15 3.55 2.1-1.2-2.1-3.45H15L4 3.5Z" fill="currentColor" />
              </svg>
            </ToolIcon>
          </button>
          <button
            aria-label="Create bubble note"
            className="annotation-tool-button annotation-tool-button-icon"
            data-active={activeTool === "bubble"}
            onClick={() => onToolChange("bubble")}
            title="Bubble note"
            type="button"
          >
            <ToolIcon>
              <svg fill="none" viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg">
                <path
                  d="M4 5.75A2.75 2.75 0 0 1 6.75 3h6.5A2.75 2.75 0 0 1 16 5.75v4.5A2.75 2.75 0 0 1 13.25 13h-4.6L5.5 15.75V13H6.75A2.75 2.75 0 0 1 4 10.25v-4.5Z"
                  stroke="currentColor"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  strokeWidth="1.6"
                />
              </svg>
            </ToolIcon>
          </button>
          <button
            aria-label="Create text box"
            className="annotation-tool-button annotation-tool-button-icon"
            data-active={activeTool === "text"}
            onClick={() => onToolChange("text")}
            title="Text box"
            type="button"
          >
            <ToolIcon>
              <svg fill="none" viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg">
                <path d="M5 5.25h10M10 5.25v9.5M7.5 14.75h5" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.7" />
              </svg>
            </ToolIcon>
          </button>
        </div>

        {selectedAnnotation ? (
          <>
            <div className="annotation-toolbar-divider" />

            <div className="annotation-toolbar-group">
              <button
                aria-label="Toggle bold"
                aria-pressed={selectedAnnotation.style.bold}
                className="annotation-style-button annotation-style-button-icon"
                onClick={() =>
                  onUpdateSelected((annotation) => ({
                    ...annotation,
                    style: {
                      ...annotation.style,
                      bold: !annotation.style.bold,
                    },
                  }))
                }
                title="Bold"
                type="button"
              >
                B
              </button>
              <button
                aria-label="Toggle italic"
                aria-pressed={selectedAnnotation.style.italic}
                className="annotation-style-button annotation-style-button-icon"
                onClick={() =>
                  onUpdateSelected((annotation) => ({
                    ...annotation,
                    style: {
                      ...annotation.style,
                      italic: !annotation.style.italic,
                    },
                  }))
                }
                title="Italic"
                type="button"
              >
                I
              </button>

              <MenuShell
                isOpen={openMenu === "fontSize"}
                menu={
                  <div className="annotation-menu-list">
                    {ANNOTATION_FONT_SIZES.map((size) => (
                      <button
                        className="annotation-menu-item"
                        data-active={selectedAnnotation.style.fontSize === size}
                        key={size}
                        onClick={() => {
                          onUpdateSelected((annotation) => ({
                            ...annotation,
                            style: {
                              ...annotation.style,
                              fontSize: size,
                            },
                          }));
                          setOpenMenu(null);
                        }}
                        type="button"
                      >
                        <span>{size}px</span>
                      </button>
                    ))}
                  </div>
                }
              >
                <button
                  aria-expanded={openMenu === "fontSize"}
                  className="annotation-menu-trigger"
                  onClick={() => toggleMenu("fontSize")}
                  type="button"
                >
                  <span>{selectedAnnotation.style.fontSize}px</span>
                  <ChevronIcon />
                </button>
              </MenuShell>

              <MenuShell
                isOpen={openMenu === "textColor"}
                menu={
                  <div className="annotation-swatch-menu">
                    {ANNOTATION_TEXT_COLORS.map((colour) => (
                      <button
                        aria-label={`Text colour ${colour}`}
                        className="annotation-colour-swatch"
                        data-active={selectedAnnotation.style.textColor === colour}
                        key={`text-${colour}`}
                        onClick={() => {
                          onUpdateSelected((annotation) => ({
                            ...annotation,
                            style: {
                              ...annotation.style,
                              textColor: colour,
                            },
                          }));
                          setOpenMenu(null);
                        }}
                        style={{ background: colour }}
                        type="button"
                      />
                    ))}
                  </div>
                }
              >
                <button
                  aria-expanded={openMenu === "textColor"}
                  className="annotation-menu-trigger annotation-menu-trigger-swatch"
                  onClick={() => toggleMenu("textColor")}
                  title="Text colour"
                  type="button"
                >
                  <span className="annotation-trigger-chip" style={{ background: selectedAnnotation.style.textColor }} />
                  <span>Text</span>
                  <ChevronIcon />
                </button>
              </MenuShell>

              {selectedAnnotation.type === "bubble" ? (
                <>
                  <MenuShell
                    isOpen={openMenu === "fillColor"}
                    menu={
                      <div className="annotation-swatch-menu">
                        {ANNOTATION_SURFACE_COLORS.map((colour) => (
                          <button
                            aria-label={`Bubble fill ${colour}`}
                            className="annotation-colour-swatch"
                            data-active={selectedAnnotation.style.fillColor === colour}
                            key={`fill-${colour}`}
                            onClick={() => {
                              onUpdateSelected((annotation) => ({
                                ...annotation,
                                style: {
                                  ...annotation.style,
                                  fillColor: colour,
                                },
                              }));
                              setOpenMenu(null);
                            }}
                            style={{ background: colour }}
                            type="button"
                          />
                        ))}
                      </div>
                    }
                  >
                    <button
                      aria-expanded={openMenu === "fillColor"}
                      className="annotation-menu-trigger annotation-menu-trigger-swatch"
                      onClick={() => toggleMenu("fillColor")}
                      title="Bubble fill"
                      type="button"
                    >
                      <span className="annotation-trigger-chip annotation-trigger-chip-border" style={{ background: selectedAnnotation.style.fillColor }} />
                      <span>Fill</span>
                      <ChevronIcon />
                    </button>
                  </MenuShell>

                  <MenuShell
                    isOpen={openMenu === "borderColor"}
                    menu={
                      <div className="annotation-swatch-menu">
                        {ANNOTATION_BORDER_COLORS.map((colour) => (
                          <button
                            aria-label={`Bubble border ${colour}`}
                            className="annotation-colour-swatch annotation-colour-swatch-outline"
                            data-active={selectedAnnotation.style.borderColor === colour}
                            key={`border-${colour}`}
                            onClick={() => {
                              onUpdateSelected((annotation) => ({
                                ...annotation,
                                style: {
                                  ...annotation.style,
                                  borderColor: colour,
                                },
                              }));
                              setOpenMenu(null);
                            }}
                            style={{ background: "#ffffff", color: colour }}
                            type="button"
                          >
                            <span style={{ background: colour }} />
                          </button>
                        ))}
                      </div>
                    }
                  >
                    <button
                      aria-expanded={openMenu === "borderColor"}
                      className="annotation-menu-trigger annotation-menu-trigger-swatch"
                      onClick={() => toggleMenu("borderColor")}
                      title="Bubble border"
                      type="button"
                    >
                      <span
                        className="annotation-trigger-chip annotation-trigger-chip-outline"
                        style={{ borderColor: selectedAnnotation.style.borderColor }}
                      />
                      <span>Border</span>
                      <ChevronIcon />
                    </button>
                  </MenuShell>
                </>
              ) : null}

              <button
                aria-label="Delete selected annotation"
                className="annotation-delete-button annotation-style-button-icon"
                onClick={onDeleteSelected}
                title="Delete"
                type="button"
              >
                <ToolIcon>
                  <svg fill="none" viewBox="0 0 20 20" xmlns="http://www.w3.org/2000/svg">
                    <path d="M6.5 7.25v7M10 7.25v7M13.5 7.25v7M4.75 5.25h10.5M8 5.25V3.75h4v1.5M6 16.25h8a1 1 0 0 0 1-1v-10H5v10a1 1 0 0 0 1 1Z" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" />
                  </svg>
                </ToolIcon>
              </button>
            </div>
          </>
        ) : null}

        <div className="annotation-toolbar-status" data-tone={saveState}>
          {saveMessage}
        </div>
      </div>
    </div>
  );
}
