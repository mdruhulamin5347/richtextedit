"use client";

import * as React from "react";

export interface UploadedFile {
  key: string;
  name: string;
  size: number;
  type: string;
  url: string;
  appUrl?: string;
}

interface UseBase64UploadProps {
  onUploadComplete?: (file: UploadedFile) => void;
  onUploadError?: (error: unknown) => void;
}

/**
 * Hook for uploading files as base64 data URLs
 * Replaces uploadthing with direct base64 conversion
 */
export function useBase64Upload({ onUploadComplete, onUploadError }: UseBase64UploadProps = {}) {
  const [uploadedFile, setUploadedFile] = React.useState<UploadedFile>();
  const [uploadingFile, setUploadingFile] = React.useState<File>();
  const [progress, setProgress] = React.useState<number>(0);
  const [isUploading, setIsUploading] = React.useState(false);

  const uploadFile = React.useCallback(
    async (file: File): Promise<UploadedFile | undefined> => {
      setIsUploading(true);
      setUploadingFile(file);
      setProgress(0);

      try {
        // Simulate progress while reading file
        const progressInterval = setInterval(() => {
          setProgress((prev) => Math.min(prev + 10, 90));
        }, 50);

        // Convert file to base64
        const base64 = await fileToBase64(file);

        clearInterval(progressInterval);
        setProgress(100);

        const uploaded: UploadedFile = {
          key: `base64-${Date.now()}-${file.name}`,
          name: file.name,
          size: file.size,
          type: file.type,
          url: base64,
          appUrl: base64,
        };

        setUploadedFile(uploaded);
        onUploadComplete?.(uploaded);

        return uploaded;
      } catch (error) {
        console.error("Error converting file to base64:", error);
        onUploadError?.(error);
        return undefined;
      } finally {
        setProgress(0);
        setIsUploading(false);
        setUploadingFile(undefined);
      }
    },
    [onUploadComplete, onUploadError]
  );

  return {
    isUploading,
    progress,
    uploadedFile,
    uploadFile,
    uploadingFile,
  };
}

/**
 * Convert a File to a base64 data URL
 */
export function fileToBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      if (typeof reader.result === "string") {
        resolve(reader.result);
      } else {
        reject(new Error("Failed to read file as base64"));
      }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

/**
 * Compress an image file before converting to base64
 * Useful for reducing the size of embedded images
 */
export async function compressAndConvertToBase64(
  file: File,
  options: {
    maxWidth?: number;
    maxHeight?: number;
    quality?: number;
  } = {}
): Promise<string> {
  const { maxWidth = 1200, maxHeight = 1200, quality = 0.8 } = options;

  // Only compress images
  if (!file.type.startsWith("image/")) {
    return fileToBase64(file);
  }

  return new Promise((resolve, reject) => {
    const img = new Image();
    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");

    if (!ctx) {
      reject(new Error("Could not get canvas context"));
      return;
    }

    img.onload = () => {
      let { width, height } = img;

      // Calculate new dimensions
      if (width > maxWidth || height > maxHeight) {
        const ratio = Math.min(maxWidth / width, maxHeight / height);
        width = Math.round(width * ratio);
        height = Math.round(height * ratio);
      }

      canvas.width = width;
      canvas.height = height;
      ctx.drawImage(img, 0, 0, width, height);

      // Convert to base64
      const base64 = canvas.toDataURL(file.type, quality);
      resolve(base64);
    };

    img.onerror = () => reject(new Error("Failed to load image"));

    // Create object URL for the file
    const url = URL.createObjectURL(file);
    img.src = url;

    // Cleanup
    img.onload = function () {
      let { width, height } = img;

      if (width > maxWidth || height > maxHeight) {
        const ratio = Math.min(maxWidth / width, maxHeight / height);
        width = Math.round(width * ratio);
        height = Math.round(height * ratio);
      }

      canvas.width = width;
      canvas.height = height;
      ctx.drawImage(img, 0, 0, width, height);

      URL.revokeObjectURL(url);
      resolve(canvas.toDataURL(file.type, quality));
    };
  });
}
