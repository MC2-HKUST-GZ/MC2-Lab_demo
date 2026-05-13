import React from 'react';

interface InlineMc2TextProps {
  text: string;
}

export default function InlineMc2Text({ text }: InlineMc2TextProps): JSX.Element {
  const parts = text.split('MC2');
  return (
    <>
      {parts.map((part, index) => (
        <React.Fragment key={`${part}-${index}`}>
          {part}
          {index < parts.length - 1 ? <>MC<sup>2</sup></> : null}
        </React.Fragment>
      ))}
    </>
  );
}
