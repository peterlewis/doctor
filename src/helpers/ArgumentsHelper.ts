

export class ArgumentsHelper {

  /**
   * Parse the command string to arguments
   * @param command 
   */
  public static parse(command: string) {
    const args: string[] = [];
    let current = "";
    let quote: string | null = null;
    let escape = false;

    const pushCurrent = () => {
      if (current.length > 0) {
        args.push(current);
        current = "";
      }
    };

    for (const ch of command) {
      if (escape) {
        current += ch;
        escape = false;
        continue;
      }

      if (ch === "\\") {
        escape = true;
        continue;
      }

      if (quote) {
        if (ch === quote) {
          quote = null;
        } else {
          current += ch;
        }
        continue;
      }

      if (ch === "'" || ch === `"`) {
        quote = ch;
        continue;
      }

      if (/\s/.test(ch)) {
        pushCurrent();
        continue;
      }

      current += ch;
    }

    if (quote || current) {
      pushCurrent();
    }

    // Re-quote args that contain spaces so execAsync receives a shell-safe command.
    return args
      .filter((arg) => typeof arg === "string")
      .map((arg) =>
        /\s/.test(arg) ? `"${arg.replace(/(["$`\\])/g, "\\$1")}"` : arg
      );
  }
}
