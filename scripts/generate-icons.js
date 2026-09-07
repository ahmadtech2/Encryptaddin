const fs = require("fs");
const path = require("path");
const zlib = require("zlib");

function crc32(buf) {
  let c = ~0;
  for (let i = 0; i < buf.length; i += 1) {
    c ^= buf[i];
    for (let k = 0; k < 8; k += 1) {
      c = c & 1 ? (c >>> 1) ^ 0xedb88320 : c >>> 1;
    }
  }
  return ~c >>> 0;
}

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const body = Buffer.concat([Buffer.from(type), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([len, body, crc]);
}

function png(size) {
  const raw = [];
  const cx = (size - 1) / 2;
  const cy = (size - 1) / 2;
  const r = size * 0.36;
  for (let y = 0; y < size; y += 1) {
    raw.push(0);
    for (let x = 0; x < size; x += 1) {
      const d = Math.hypot(x - cx, y - cy * 0.92);
      const inDisc = d < r;
      const stem = Math.abs(x - cx) < size * 0.08 && y > cy && y < size * 0.78;
      const on = inDisc || stem;
      raw.push(on ? 31 : 11, on ? 138 : 31, on ? 112 : 58, 255);
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8;
  ihdr[9] = 6;
  return Buffer.concat([
    Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
    chunk("IHDR", ihdr),
    chunk("IDAT", zlib.deflateSync(Buffer.from(raw))),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

const dir = path.join(__dirname, "..", "assets");
fs.mkdirSync(dir, { recursive: true });
for (const size of [16, 32, 64, 80]) {
  fs.writeFileSync(path.join(dir, `icon-${size}.png`), png(size));
}
