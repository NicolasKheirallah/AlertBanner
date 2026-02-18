/**
 * AlertListShimmer Component
 * 
 * A loading state component for the alert list using Fluent UI Shimmer.
 * Displays 3 shimmer rows that match the approximate height of alert items.
 * 
 * @example
 * {isLoadingAlerts && <AlertListShimmer />}
 */

import * as React from "react";
import {
  Shimmer,
  ShimmerElementType,
  IShimmerStyleProps,
  IShimmerStyles,
} from "@fluentui/react";

/**
 * Shimmer styles override for consistent appearance
 */
const shimmerStyles = (props: IShimmerStyleProps): IShimmerStyles => ({
  root: {
    padding: "12px 0",
  },
  shimmerWrapper: [
    {
      backgroundColor: "#f3f2f1",
    },
    props.isDataLoaded && {
      backgroundColor: "transparent",
    },
  ],
  dataWrapper: {
    padding: "0",
  },
});

/**
 * AlertListShimmer Component
 * 
 * Renders shimmer placeholder rows while alerts are loading.
 * Each row uses a circle + gap + line pattern to simulate alert items.
 */
const AlertListShimmer: React.FC<{
  rowCount?: number;
  className?: string;
  dataTestId?: string;
}> = ({
  rowCount = 3,
  className,
  dataTestId,
}) => {
  /**
   * Creates shimmer elements for a single row
   * Pattern: Circle (avatar) + gap + Line (title) + gap + Line (description)
   */
  const getShimmerElements = () => [
    // Circle for avatar/icon area
    { type: ShimmerElementType.circle, height: 40, width: 40 },
    // Gap between circle and content
    { type: ShimmerElementType.gap, height: 40, width: 16 },
    // Group of lines for content
    {
      type: ShimmerElementType.line,
      height: 12,
      width: "40%",
    },
    // Gap between lines
    { type: ShimmerElementType.gap, height: 8, width: "100%" },
    // Second line (description)
    {
      type: ShimmerElementType.line,
      height: 10,
      width: "70%",
    },
  ];

  return (
    <div
      className={className}
      data-testid={dataTestId}
      role="progressbar"
      aria-label="Loading alerts..."
      aria-busy="true"
    >
      {Array.from({ length: rowCount }).map((_, index) => (
        <Shimmer
          key={index}
          shimmerElements={getShimmerElements()}
          styles={shimmerStyles}
          // Add slight delay for cascading effect
          style={{ animationDelay: `${index * 100}ms` }}
        />
      ))}
    </div>
  );
};

export default AlertListShimmer;
