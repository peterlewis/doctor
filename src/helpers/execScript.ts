import { CliCommand } from "./CliCommand";
import { spawn } from "cross-spawn";
import { Logger } from "./logger";
import { defer, StatusHelper } from ".";
import { Deferred } from "@models";
import { execAsync } from "@utils";

// Catch all the errors coming from command execution
process.on("unhandledRejection", (reason, promise) => {
  Logger.debug(`Unhandled Rejection at: ${reason}`);
});

/**
 * Execute script with retry logic
 * @param args
 * @param shouldRetry
 * @param shouldSpawn
 * @param toMask
 */
export const execScript = async <T>(
  args: string[] = [],
  shouldRetry: boolean = false,
  shouldSpawn: boolean = false,
  toMask: string[] = [],
  deferred?: Deferred<T>
): Promise<T | any> => {
  let firstRun = false;
  if (!deferred) {
    firstRun = true;
    deferred = defer();
  }

  promiseExecScript<T>(args, shouldSpawn, toMask)
    .then((result) => {
      deferred.resolve(result);
    })
    .catch((err) => {
      if (shouldRetry && firstRun) {
        Logger.debug(`Doctor will retry to execute the command again.`);
        StatusHelper.addRetry();

        // Waiting 5 seconds in order to make sure that the call did not happen too fast after the previous failure
        setTimeout(async () => {
          execScript(args, shouldRetry, shouldSpawn, toMask, deferred);
        }, 5000);
      } else {
        deferred.reject(err);
      }
    });

  return deferred.promise;
};

const promiseExecScript = async <T>(
  args: string[] = [],
  shouldSpawn: boolean = false,
  toMask: string[] = []
): Promise<T> => {
  return new Promise<T>(async (resolve, reject) => {
    Logger.debug(``);
    const cmdToExec = Logger.mask(
      `${CliCommand.getName()} ${args.join(" ")}`,
      toMask
    );
    const startTime = Date.now();
    const logPhase = (phase: "start" | "success" | "error", extra?: string) => {
      const duration = phase === "start" ? 0 : Date.now() - startTime;
      const durationMsg = phase === "start" ? "" : ` (${duration}ms)`;
      const suffix = extra ? ` ${extra}` : "";
      Logger.debug(`[exec:${phase}] ${cmdToExec}${durationMsg}${suffix}`);
    };
    logPhase("start");

    if (CliCommand.getDryRun()) {
      Logger.debug(`[dry-run] Skipping execution: ${cmdToExec}`);
      return resolve(("" as unknown) as T);
    }

    if (shouldSpawn) {
      const execution = spawn(CliCommand.getName(), [...args]);
      let finished = false;

      execution.stdout.on("data", (data) => {
        console.log(`${data}`);
      });

      execution.stdout.on("close", (data: any) => {
        if (finished) {
          return;
        }
        finished = true;
        logPhase("success");
        resolve(data);
      });

      execution.stderr.on("data", async (error) => {
        if (finished) {
          return;
        }
        finished = true;
        const maskedError = Logger.mask(
          error ? error.toString() : "",
          toMask
        );
        logPhase("error", maskedError);
        reject(new Error(maskedError));
      });
    } else {
      try {
        const { stdout, stderr } = await execAsync(
          `${CliCommand.getName()} ${args.join(" ")}`
        );
        if (stderr) {
          const error = Logger.mask(stderr.toString(), toMask);
          logPhase("error", error);
          reject(new Error(error));
          return;
        }

        logPhase("success");
        resolve(stdout as any as T);
      } catch (e) {
        const error = e instanceof Error ? e : new Error(e as any);
        logPhase("error", error.message);
        reject(error);
      }
    }
  });
};
