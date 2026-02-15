import 'dotenv/config';
import { Client } from 'minio';

const minioClient = new Client({
  endPoint: process.env.MINIO_ENDPOINT || 'minio',
  port: Number(process.env.MINIO_PORT || 9000),
  useSSL: process.env.MINIO_USE_SSL === 'true',
  accessKey: process.env.MINIO_ACCESS_KEY,
  secretKey: process.env.MINIO_SECRET_KEY,
});


export async function ensureBucket() {
  const exists = await minioClient.bucketExists(RECOVERY_BUCKET).catch(() => false);
  if (!exists) {
    await minioClient.makeBucket(RECOVERY_BUCKET, '');
  }
}

export { minioClient };
