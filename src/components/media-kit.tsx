"use client";

import { CaptionPlugin } from "@platejs/caption/react";
import {
  ImagePlugin,
  PlaceholderPlugin,
} from "@platejs/media/react";
import { KEYS } from "platejs";

import { ImageElement } from "@/components/ui/media-image-node";
import { PlaceholderElement } from "@/components/ui/media-placeholder-node";
import { MediaPreviewDialog } from "@/components/ui/media-preview-dialog";
import { MediaUploadToast } from "@/components/ui/media-upload-toast";

// Deviation from the upstream editor: only IMAGES are enabled. The embed / video / audio
// / file plugins are deliberately absent — none of them can survive
// serialization to the HTML that dompdf and the print blades consume, and
// between them they dragged react-player, dash.js and hls.js (~1.5MB) into the
// bundle for a lab report. The serializer still degrades any such node found in
// legacy content to a link rather than dropping it.
export const MediaKit = [
  ImagePlugin.configure({
    options: { disableUploadInsert: true },
    render: { afterEditable: MediaPreviewDialog, node: ImageElement },
  }),
  PlaceholderPlugin.configure({
    options: { disableEmptyPlaceholder: true },
    render: { afterEditable: MediaUploadToast, node: PlaceholderElement },
  }),
  CaptionPlugin.configure({
    options: {
      query: {
        allow: [KEYS.img],
      },
    },
  }),
];
