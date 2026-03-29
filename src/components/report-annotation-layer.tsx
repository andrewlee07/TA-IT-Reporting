"use client";

import { useEffect, useMemo, useRef, useState, type PointerEvent as ReactPointerEvent } from "react";

import type { ReportAnnotation, ReportAnnotationPoint, ReportAnnotationTool } from "@/lib/annotations/types";

const MIN_ANNOTATION_WIDTH = 0.12;
const MIN_ANNOTATION_HEIGHT = 0.08;

type DragState =
  | {
      mode: "move";
      annotationId: string;
      originClientX: number;
      originClientY: number;
      startAnnotation: ReportAnnotation;
      overlayWidth: number;
      overlayHeight: number;
    }
  | {
      mode: "resize";
      annotationId: string;
      originClientX: number;
      originClientY: number;
      startAnnotation: ReportAnnotation;
      overlayWidth: number;
      overlayHeight: number;
    }
  | {
      mode: "tail";
      annotationId: string;
      originClientX: number;
      originClientY: number;
      startAnnotation: ReportAnnotation;
      overlayWidth: number;
      overlayHeight: number;
    };

interface ReportAnnotationLayerProps {
  slideId: string;
  annotations: ReportAnnotation[];
  activeTool: ReportAnnotationTool;
  canEdit: boolean;
  selectedAnnotationId: string | null;
  onCreate: (slideId: string, type: "bubble" | "text", point: ReportAnnotationPoint) => void;
  onDelete: (annotationId: string) => void;
  onSelect: (annotationId: string | null) => void;
  onUpdate: (annotationId: string, updater: (annotation: ReportAnnotation) => ReportAnnotation) => void;
}

function clampUnit(value: number): number {
  return Math.min(1, Math.max(0, value));
}

function renderText(value: string): string[] {
  return value.split("\n");
}

function getBubbleAttachPoint(annotation: ReportAnnotation): ReportAnnotationPoint | null {
  if (!annotation.tailAnchor) {
    return null;
  }

  const centerX = annotation.x + annotation.width / 2;
  const centerY = annotation.y + annotation.height / 2;
  const deltaX = annotation.tailAnchor.x - centerX;
  const deltaY = annotation.tailAnchor.y - centerY;

  if (deltaX === 0 && deltaY === 0) {
    return {
      x: centerX,
      y: annotation.y + annotation.height,
    };
  }

  const scaleX = deltaX === 0 ? Number.POSITIVE_INFINITY : (deltaX > 0 ? annotation.width / 2 : -annotation.width / 2) / deltaX;
  const scaleY = deltaY === 0 ? Number.POSITIVE_INFINITY : (deltaY > 0 ? annotation.height / 2 : -annotation.height / 2) / deltaY;
  const scale = Math.min(Math.abs(scaleX), Math.abs(scaleY));

  return {
    x: centerX + deltaX * scale,
    y: centerY + deltaY * scale,
  };
}

function getBubbleTailPolygon(annotation: ReportAnnotation): string | null {
  if (annotation.type !== "bubble" || !annotation.tailAnchor) {
    return null;
  }

  const attachPoint = getBubbleAttachPoint(annotation);
  if (!attachPoint) {
    return null;
  }

  const tolerance = 0.0005;
  const onTopOrBottom =
    Math.abs(attachPoint.y - annotation.y) < tolerance || Math.abs(attachPoint.y - (annotation.y + annotation.height)) < tolerance;
  const baseOffset = onTopOrBottom
    ? { x: Math.min(annotation.width * 0.12, 0.022), y: 0 }
    : { x: 0, y: Math.min(annotation.height * 0.16, 0.022) };

  const pointA = {
    x: attachPoint.x - baseOffset.x,
    y: attachPoint.y - baseOffset.y,
  };
  const pointB = {
    x: attachPoint.x + baseOffset.x,
    y: attachPoint.y + baseOffset.y,
  };

  return `${annotation.tailAnchor.x * 100},${annotation.tailAnchor.y * 100} ${pointA.x * 100},${pointA.y * 100} ${pointB.x * 100},${pointB.y * 100}`;
}

export function ReportAnnotationLayer({
  slideId,
  annotations,
  activeTool,
  canEdit,
  selectedAnnotationId,
  onCreate,
  onDelete,
  onSelect,
  onUpdate,
}: ReportAnnotationLayerProps) {
  const overlayRef = useRef<HTMLDivElement | null>(null);
  const [dragState, setDragState] = useState<DragState | null>(null);
  const slideAnnotations = useMemo(
    () =>
      annotations
        .filter((annotation) => annotation.slideId === slideId)
        .sort((left, right) => left.zIndex - right.zIndex || left.id.localeCompare(right.id)),
    [annotations, slideId],
  );

  useEffect(() => {
    if (!dragState) {
      return;
    }

    const currentDragState = dragState;

    function handlePointerMove(event: PointerEvent) {
      const deltaX = (event.clientX - currentDragState.originClientX) / currentDragState.overlayWidth;
      const deltaY = (event.clientY - currentDragState.originClientY) / currentDragState.overlayHeight;

      if (currentDragState.mode === "move") {
        onUpdate(currentDragState.annotationId, () => ({
          ...currentDragState.startAnnotation,
          x: clampUnit(Math.min(currentDragState.startAnnotation.x + deltaX, 1 - currentDragState.startAnnotation.width)),
          y: clampUnit(Math.min(currentDragState.startAnnotation.y + deltaY, 1 - currentDragState.startAnnotation.height)),
        }));
        return;
      }

      if (currentDragState.mode === "resize") {
        onUpdate(currentDragState.annotationId, () => ({
          ...currentDragState.startAnnotation,
          width: clampUnit(Math.max(MIN_ANNOTATION_WIDTH, Math.min(currentDragState.startAnnotation.width + deltaX, 1 - currentDragState.startAnnotation.x))),
          height: clampUnit(Math.max(MIN_ANNOTATION_HEIGHT, Math.min(currentDragState.startAnnotation.height + deltaY, 1 - currentDragState.startAnnotation.y))),
        }));
        return;
      }

      if (currentDragState.startAnnotation.tailAnchor) {
        onUpdate(currentDragState.annotationId, () => ({
          ...currentDragState.startAnnotation,
          tailAnchor: {
            x: clampUnit(currentDragState.startAnnotation.tailAnchor!.x + deltaX),
            y: clampUnit(currentDragState.startAnnotation.tailAnchor!.y + deltaY),
          },
        }));
      }
    }

    function handlePointerUp() {
      setDragState(null);
    }

    window.addEventListener("pointermove", handlePointerMove);
    window.addEventListener("pointerup", handlePointerUp);

    return () => {
      window.removeEventListener("pointermove", handlePointerMove);
      window.removeEventListener("pointerup", handlePointerUp);
    };
  }, [dragState, onUpdate]);

  function toRelativePoint(clientX: number, clientY: number): ReportAnnotationPoint | null {
    const overlay = overlayRef.current;
    if (!overlay) {
      return null;
    }

    const rect = overlay.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0) {
      return null;
    }

    return {
      x: clampUnit((clientX - rect.left) / rect.width),
      y: clampUnit((clientY - rect.top) / rect.height),
    };
  }

  function startDrag(event: ReactPointerEvent<HTMLButtonElement>, mode: DragState["mode"], annotation: ReportAnnotation) {
    if (!canEdit || !overlayRef.current) {
      return;
    }

    const rect = overlayRef.current.getBoundingClientRect();
    event.preventDefault();
    event.stopPropagation();
    onSelect(annotation.id);
    setDragState({
      mode,
      annotationId: annotation.id,
      originClientX: event.clientX,
      originClientY: event.clientY,
      startAnnotation: annotation,
      overlayWidth: rect.width || 1,
      overlayHeight: rect.height || 1,
    });
  }

  function handleStagePointerDown(event: ReactPointerEvent<HTMLDivElement>) {
    if (!canEdit) {
      return;
    }

    if (activeTool === "select") {
      onSelect(null);
      return;
    }

    const point = toRelativePoint(event.clientX, event.clientY);
    if (!point) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();
    onCreate(slideId, activeTool, point);
  }

  const shouldCaptureBackground = canEdit && (activeTool !== "select" || selectedAnnotationId !== null);

  return (
    <div className="annotation-layer annotation-layer-live" ref={overlayRef}>
      <svg aria-hidden="true" className="annotation-tail-stage" preserveAspectRatio="none" viewBox="0 0 100 100">
        {slideAnnotations.map((annotation) => {
          const polygon = getBubbleTailPolygon(annotation);
          if (!polygon) {
            return null;
          }

          return (
            <polygon
              className="annotation-tail"
              key={`tail-${annotation.id}`}
              points={polygon}
              style={{
                fill: annotation.style.fillColor,
                stroke: annotation.style.borderColor,
              }}
            />
          );
        })}
      </svg>

      {shouldCaptureBackground ? (
        <div
          aria-hidden="true"
          className="annotation-stage-capture"
          data-mode={activeTool !== "select" ? "create" : "select"}
          onPointerDown={handleStagePointerDown}
        />
      ) : null}

      {slideAnnotations.map((annotation) => {
        const isSelected = annotation.id === selectedAnnotationId;
        const lines = renderText(annotation.text || "");

        return (
          <div
            className="annotation-note-shell"
            key={annotation.id}
            style={{
              left: `${annotation.x * 100}%`,
              top: `${annotation.y * 100}%`,
              width: `${annotation.width * 100}%`,
              height: `${annotation.height * 100}%`,
              zIndex: 20 + annotation.zIndex,
            }}
          >
            {canEdit && isSelected ? (
              <button
                className="annotation-drag-handle"
                onPointerDown={(event) => startDrag(event, "move", annotation)}
                type="button"
              >
                Move
              </button>
            ) : null}

            <div
              className={`annotation-note annotation-note-${annotation.type}${isSelected ? " is-selected" : ""}`}
              onPointerDown={(event) => {
                event.stopPropagation();
                onSelect(annotation.id);
              }}
              style={{
                color: annotation.style.textColor,
                background: annotation.style.fillColor,
                borderColor: annotation.style.borderColor,
                fontSize: `${annotation.style.fontSize}px`,
                fontWeight: annotation.style.bold ? 700 : 500,
                fontStyle: annotation.style.italic ? "italic" : "normal",
              }}
            >
              {canEdit && isSelected ? (
                <textarea
                  autoFocus
                  aria-label={annotation.type === "bubble" ? "Bubble note text" : "Text box content"}
                  className="annotation-note-input"
                  onClick={(event) => event.stopPropagation()}
                  onChange={(event) =>
                    onUpdate(annotation.id, (current) => ({
                      ...current,
                      text: event.target.value,
                    }))
                  }
                  value={annotation.text}
                />
              ) : (
                <div className="annotation-note-content">
                  {lines.map((line, index) => (
                    <span key={`${annotation.id}-line-${index}`}>
                      {line || "\u00A0"}
                      {index < lines.length - 1 ? <br /> : null}
                    </span>
                  ))}
                </div>
              )}
            </div>

            {canEdit && isSelected ? (
              <button
                className="annotation-resize-handle"
                onPointerDown={(event) => startDrag(event, "resize", annotation)}
                type="button"
              />
            ) : null}

            {canEdit && isSelected ? (
              <button
                className="annotation-delete-pill"
                onClick={() => onDelete(annotation.id)}
                type="button"
              >
                Delete
              </button>
            ) : null}
          </div>
        );
      })}

      {canEdit
        ? slideAnnotations.map((annotation) =>
            annotation.id === selectedAnnotationId && annotation.type === "bubble" && annotation.tailAnchor ? (
              <button
                className="annotation-tail-handle"
                key={`tail-handle-${annotation.id}`}
                onPointerDown={(event) => startDrag(event, "tail", annotation)}
                style={{
                  left: `${annotation.tailAnchor.x * 100}%`,
                  top: `${annotation.tailAnchor.y * 100}%`,
                }}
                type="button"
              />
            ) : null,
          )
        : null}
    </div>
  );
}
