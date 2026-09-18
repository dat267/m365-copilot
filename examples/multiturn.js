// Example: one conversation, two turns, streamed output.
//   M365_INSECURE=1 node examples/multiturn.js
import { M365Session, ask } from "../src/index.js";

const session = new M365Session();

// Turn 1 (streamed)
const stream = await session.chat("My favorite color is teal. Acknowledge in one word.");
for await (const delta of stream) process.stdout.write(delta);
process.stdout.write("\n");

// Turn 2, same conversation — M365 should remember.
console.log("recall:", await ask("What is my favorite color? One word.", { session }));
