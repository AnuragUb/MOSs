# Media Operations Suite (MOS) - User Guide

Welcome to the MOS application! This guide will walk you through the primary features and how to use them effectively.

## Table of Contents
1.  [Getting Started: Loading a Video](#1-getting-started-loading-a-video)
2.  [The Interface](#2-the-interface)
3.  [Creating Markers](#3-creating-markers)
    *   [Single Button Mode](#single-button-mode)
    *   [TCR In / TCR Out Mode](#tcr-in--tcr-out-mode)
4.  [Editing Marker Data](#4-editing-marker-data)
5.  [Audio Recognition](#5-audio-recognition)
6.  [Marking Rows](#6-marking-rows)
    *   [Marking for Emphasis](#marking-for-emphasis)
    *   [Marking for Exceptions](#marking-for-exceptions)
7.  [Exporting Data](#7-exporting-data)
8.  [Managing Music Companies](#8-managing-music-companies)

---

### 1. Getting Started: Loading a Video

You have two options for loading a video:

*   **Load Local Video**: Click the **"Load Local Video"** button and select a video file from your computer.
*   **Enter Video URL**: Paste a direct video URL (e.g., from a cloud storage link) into the text box and click **"Load URL"**. The application will load the video from the provided link.

### 2. The Interface

The screen is divided into two main sections:
*   **Left Side**: The video player and its controls.
*   **Right Side**: The marker table where your created timestamps and metadata will appear.

You can resize these sections by clicking and dragging the border between them.

### 3. Creating Markers

Markers are timestamped rows in the table that correspond to a segment of the video. There are two primary modes for creating them, selectable via the radio buttons on the left.

#### Single Button Mode
This is the fastest way to create sequential markers.
1.  Select **"Single Button Mode"**.
2.  Play the video.
3.  Press the **"TCR In"** button at the exact moment a segment begins. This action:
    *   Sets the "TCR Out" of the *previous* marker to the current time.
    *   Creates a *new* marker, setting its "TCR In" to the current time.
    *   Calculates the duration for the previous marker.
4.  When you are finished with the last segment, press the **"TCR Out"** button to finalize the end time for the very last marker.

#### TCR In / TCR Out Mode
This mode gives you more deliberate control over each segment.
1.  Select **"TCR In / TCR Out Mode"**.
2.  Play the video.
3.  At the start of a segment, click the **"TCR In"** button.
4.  At the end of the segment, click the **"TCR Out"** button.
5.  This creates a single, complete marker with its In, Out, and Duration times filled.

### 4. Editing Marker Data

The table on the right is fully editable.
*   **Click on any cell** (like "Title", "Film/Album", etc.) to type directly into it.
*   Changes are saved automatically as you move to another cell.
*   Use the **"Add Column"** button to add custom columns to your table for any additional data you need to track.

### 5. Audio Recognition

For any marker, you can identify the music playing in that segment.
1.  In the row for the desired marker, click the **"Recognize"** button (looks like a magnifying glass).
2.  The system will process the audio for that segment and automatically fill in the `Title`, `Artist`, and other related music information if a match is found on AudD.

### 6. Marking Rows

You can highlight specific rows for attention or to denote special conditions.

#### Marking for Emphasis
To highlight a row:
1.  **Double-click** the checkbox cell (the very first cell) of a row to mark it **Yellow**.
2.  **Triple-click** the checkbox cell to mark it **Red**.
3.  Clicking a fourth time will clear the emphasis mark.

These colors will be reflected in the final Excel export.

#### Marking for Exceptions
To flag a row as an exception (e.g., for review):
1.  Click the **"Mark Exceptions"** button. A dialog will appear.
2.  Select the reason for the exception (e.g., "Title column exception").
3.  Click **"Apply Exception"**. The selected row will be marked with a grey background and a red dot.
4.  To remove an exception, click the small red 'x' button that appears next to the checkbox.

### 7. Exporting Data

When your work is complete, you can export the data to an Excel or CSV file.
1.  Click the **"Export Settings"** button.
2.  In the settings page, you can:
    *   Choose which columns to include in the export.
    *   Select the time format (`HH:MM:SS` or `HH:MM:SS:FF`).
    *   Add custom header rows to the exported file.
3.  Once configured, return to the main page.
4.  Click the **"Export"** button. Your file will be generated and downloaded.

### 8. Managing Music Companies

The application has a separate page for managing a central database of Music Companies.
*   Navigate to the **/music-co** page.
*   Here you can **Add**, **Edit**, and **Delete** entries for music companies.
*   This database is used to provide consistent suggestions in the "Music Co." column on the main marker page. 