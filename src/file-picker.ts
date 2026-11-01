function pickFiles(directory: boolean): Promise<File[]> {
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept =
      ".docx,.docm,.doc,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/vnd.ms-word.document.macroEnabled.12";
    input.multiple = true;
    if (directory) {
      input.setAttribute("webkitdirectory", "");
      input.setAttribute("directory", "");
    }
    input.classList.add("sun-daoru-file-input");
    document.body.appendChild(input);

    let settled = false;
    const finish = (files: File[]) => {
      if (settled) return;
      settled = true;
      input.remove();
      resolve(files);
    };

    input.addEventListener("change", () => {
      finish(Array.from(input.files ?? []));
    });
    input.addEventListener("cancel", () => finish([]));

    window.addEventListener(
      "focus",
      () => {
        window.setTimeout(() => {
          if (!settled && !input.files?.length) finish([]);
        }, 350);
      },
      { once: true },
    );

    input.click();
  });
}

export function pickDocxFiles(): Promise<File[]> {
  return pickFiles(false);
}

export function pickDocxFolder(): Promise<File[]> {
  return pickFiles(true);
}
