
import * as React from "react";
import {
  Shimmer,
  ShimmerElementType,
  IShimmerStyleProps,
  IShimmerStyles,
} from "@fluentui/react";

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

const AlertListShimmer: React.FC<{
  rowCount?: number;
  className?: string;
  dataTestId?: string;
}> = ({
  rowCount = 3,
  className,
  dataTestId,
}) => {
    const getShimmerElements = () => [
    { type: ShimmerElementType.circle, height: 40, width: 40 },
    { type: ShimmerElementType.gap, height: 40, width: 16 },
    {
      type: ShimmerElementType.line,
      height: 12,
      width: "40%",
    },
    { type: ShimmerElementType.gap, height: 8, width: "100%" },
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
          style={{ animationDelay: `${index * 100}ms` }}
        />
      ))}
    </div>
  );
};

export default AlertListShimmer;
