'use client';

import { useRef, useState } from 'react';
import { cn } from '@/lib/utils';

export type PdfRegion = {
  id: string;
  page: number;
  type: 'text' | 'checkbox' | 'table_cell' | string;
  box: { x: number; y: number; w: number; h: number };
  value: string;
  acroFieldName?: string | null;
};

export type PdfPage = {
  index: number;
  width_pt: number;
  height_pt: number;
  image_url: string;
  regions: PdfRegion[];
};

function RegionOverlay({
  region,
  pageWidth,
  pageHeight,
  onChange,
}: {
  region: PdfRegion;
  pageWidth: number;
  pageHeight: number;
  onChange: (value: string) => void;
}) {
  const [focused, setFocused] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const style: React.CSSProperties = {
    position: 'absolute',
    left: `${(region.box.x / pageWidth) * 100}%`,
    top: `${(region.box.y / pageHeight) * 100}%`,
    width: `${(region.box.w / pageWidth) * 100}%`,
    height: `${(region.box.h / pageHeight) * 100}%`,
  };

  const fontSize = `${Math.min((region.box.h / pageHeight) * 100 * 0.6, 1.8)}cqh`;
  const hasValue = Boolean(region.value);

  if (region.type === 'checkbox') {
    return (
      <button
        style={style}
        className="flex items-center justify-center hover:bg-sky-200/50 active:bg-sky-300/50 rounded transition-colors touch-manipulation cursor-pointer"
        onClick={() => onChange(region.value === 'true' ? '' : 'true')}
        aria-label="Toggle checkbox"
      >
        {region.value === 'true' && (
          <span className="text-black font-bold leading-none select-none" style={{ fontSize }}>
            ☑
          </span>
        )}
      </button>
    );
  }

  return (
    <div style={style} className="group">
      {focused ? (
        <input
          ref={inputRef}
          autoFocus
          type="text"
          value={region.value}
          onChange={(e) => onChange(e.target.value)}
          onBlur={() => setFocused(false)}
          onKeyDown={(e) => { if (e.key === 'Enter') { e.preventDefault(); (e.target as HTMLInputElement).blur(); } }}
          className="absolute inset-0 w-full h-full bg-sky-50/95 border border-sky-400 rounded-sm px-1 text-black focus:outline-none"
          style={{ fontSize }}
        />
      ) : (
        <button
          className={cn(
            'absolute inset-0 w-full h-full text-left px-1 rounded-sm transition-colors touch-manipulation overflow-hidden',
            hasValue
              ? 'bg-sky-100/70 hover:bg-sky-200/70'
              : 'bg-transparent border border-transparent group-hover:border-sky-300/80 hover:bg-sky-50/60',
          )}
          onClick={() => {
            setFocused(true);
            setTimeout(() => inputRef.current?.focus(), 0);
          }}
          style={{ fontSize }}
        >
          {hasValue ? (
            <span className="text-black truncate block leading-tight">{region.value}</span>
          ) : (
            <span className="text-sky-400/0 group-hover:text-sky-400/60 text-[10px] transition-colors">tap</span>
          )}
        </button>
      )}
    </div>
  );
}

export function PdfFillView({
  pages,
  onRegionChange,
}: {
  pages: PdfPage[];
  onRegionChange: (pageIndex: number, regionId: string, value: string) => void;
}) {
  if (pages.length === 0) {
    return (
      <div className="flex items-center justify-center h-full text-sm text-muted-foreground">
        No pages found.
      </div>
    );
  }

  return (
    <div className="flex flex-col gap-6 p-4 overflow-y-auto h-full bg-gray-100">
      {pages.map((page) => (
        <div
          key={page.index}
          className="mx-auto shadow-md"
          style={{
            width: '100%',
            maxWidth: '800px',
            position: 'relative',
            paddingBottom: `${(page.height_pt / page.width_pt) * 100}%`,
          }}
        >
          {/* Inner div fills the padded space; containerType:'size' needs explicit dims */}
          <div
            className="absolute inset-0 bg-white"
            style={{ containerType: 'size' }}
          >
            {/* eslint-disable-next-line @next/next/no-img-element */}
            <img
              src={page.image_url}
              alt={`Page ${page.index + 1}`}
              className="absolute inset-0 w-full h-full object-fill pointer-events-none select-none"
              draggable={false}
            />

            {page.regions.map((region) => (
              <RegionOverlay
                key={region.id}
                region={region}
                pageWidth={page.width_pt}
                pageHeight={page.height_pt}
                onChange={(value) => onRegionChange(page.index, region.id, value)}
              />
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}
