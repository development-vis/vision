'use client';

import dynamic from 'next/dynamic';
import config from '../../../sanity.config'; // Ensure this path matches your setup

// This dynamically imports the studio and strictly disables Server-Side Rendering
const NextStudio = dynamic(
  () => import('next-sanity/studio').then((mod) => mod.NextStudio),
  { ssr: false }
);

export default function StudioPage() {
  return <NextStudio config={config} />
}