'use strict';
// Optional ffprobe check. Everything is best-effort: if ffprobe is missing we
// only warn, because Instagram itself is the final judge of the file.
const { execFile } = require('node:child_process');

const MAX_BYTES = 1024 * 1024 * 1024; // 1 GB for Reels
const MAX_SECONDS = 15 * 60;
const MIN_SECONDS = 3;

function runFfprobe(command, file, timeoutMs = 15_000) {
  return new Promise((resolve) => {
    execFile(
      command,
      ['-v', 'error', '-print_format', 'json', '-show_format', '-show_streams', file],
      { timeout: timeoutMs, maxBuffer: 8 * 1024 * 1024 },
      (error, stdout) => {
        if (error) return resolve(null);
        try {
          resolve(JSON.parse(stdout));
        } catch {
          resolve(null);
        }
      }
    );
  });
}

async function probeVideo(file, { command = 'ffprobe' } = {}) {
  const data = await runFfprobe(command, file);
  if (!data) return null;
  const video = (data.streams || []).find((stream) => stream.codec_type === 'video') || {};
  const duration = Number(data.format?.duration ?? video.duration);
  return {
    durationSec: Number.isFinite(duration) ? duration : null,
    width: video.width ?? null,
    height: video.height ?? null,
    codec: video.codec_name ?? null,
    audioCodec: ((data.streams || []).find((stream) => stream.codec_type === 'audio') || {}).codec_name ?? null,
  };
}

// Returns a list of warnings; an empty list means "looks fine".
function checkReelConstraints(file, info) {
  const warnings = [];
  if (file.size > MAX_BYTES) warnings.push(`حجم فایل ${(file.size / 1024 ** 3).toFixed(2)} گیگابایت است؛ سقف ریلز ۱ گیگابایت.`);
  if (!info) {
    warnings.push('ffprobe در دسترس نبود؛ مدت/کدک بررسی نشد (اختیاری، فایل باز هم ارسال می‌شود).');
    return warnings;
  }
  if (info.durationSec !== null && info.durationSec > MAX_SECONDS) {
    warnings.push(`مدت ویدیو ${Math.round(info.durationSec / 60)} دقیقه است؛ سقف ریلز ۱۵ دقیقه.`);
  }
  if (info.durationSec !== null && info.durationSec < MIN_SECONDS) {
    warnings.push(`مدت ویدیو ${info.durationSec.toFixed(1)} ثانیه است؛ حداقل ۳ ثانیه لازم است.`);
  }
  if (info.codec && !/^(h264|hevc)$/.test(info.codec)) {
    warnings.push(`کدک ویدیو ${info.codec} است؛ H.264 توصیه می‌شود.`);
  }
  if (info.audioCodec && !/^(aac|mp3)$/.test(info.audioCodec)) {
    warnings.push(`کدک صدا ${info.audioCodec} است؛ AAC توصیه می‌شود.`);
  }
  if (info.width && info.height) {
    const ratio = info.width / info.height;
    if (Math.abs(ratio - 9 / 16) > 0.06) {
      warnings.push(`نسبت تصویر ${info.width}×${info.height} عمودی ۹:۱۶ نیست؛ ریلز ممکن است برش بخورد.`);
    }
  }
  return warnings;
}

module.exports = { probeVideo, checkReelConstraints, MAX_BYTES, MAX_SECONDS };
