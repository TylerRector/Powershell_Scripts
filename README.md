\documentclass[11pt]{article}
\usepackage[margin=1in]{geometry}
\usepackage{hyperref}
\usepackage{titlesec}

\title{PowerShell Automation Scripts}
\date{}

\begin{document}

\maketitle

\section*{Overview}
Collection of PowerShell scripts for batch printing Microsoft Office files and organizing imaging data by channel.

\section*{Requirements}
\begin{itemize}
    \item Windows OS
    \item PowerShell 5+
    \item Microsoft Word (for DOCX script)
    \item Microsoft PowerPoint (for PPTX script)
\end{itemize}

\section*{Scripts}

\subsection*{1. DOCX Script}
Prints all \texttt{.docx} files in a specified folder using Microsoft Word COM automation.

\textbf{Usage:}
\begin{itemize}
    \item Set \texttt{\$folderPath} to target directory
    \item Runs silent (Word hidden by default)
\end{itemize}

\subsection*{2. PPTX Script}
Prints all \texttt{.pptx} files in a specified folder using PowerPoint COM automation.

\textbf{Usage:}
\begin{itemize}
    \item Set \texttt{\$folderPath}
    \item PowerPoint runs visible by default
\end{itemize}

\subsection*{3. Improved Multi File Script}
Processes experiment folders using pattern matching.

\textbf{Configuration:}
\begin{itemize}
    \item \texttt{\$root} $\rightarrow$ root directory
    \item \texttt{\$folderNamePattern} $\rightarrow$ regex for target folders
    \item \texttt{\$channelPattern} $\rightarrow$ regex for channel IDs
\end{itemize}

\section*{Notes}
\begin{itemize}
    \item All scripts rely on filename pattern matching (\texttt{ch\textbackslash d\{2\}})
    \item Existing files may be overwritten if names conflict
    \item Ensure paths are correctly set before execution
\end{itemize}

\end{document}
