import * as React from "react";
import { Button, Text, tokens } from "@fluentui/react-components";
import { htmlSanitizer } from "../Utils/HtmlSanitizer";
import richMediaStyles from "../Services/RichMediaAlert.module.scss";

interface IDescriptionContentProps {
  description: string;
}

const DescriptionContent: React.FC<IDescriptionContentProps> = React.memo(({ description }) => {
  const [isExpanded, setIsExpanded] = React.useState(false);
  const TRUNCATE_LENGTH = 200; // Character limit for truncation

  const toggleExpanded = () => {
    setIsExpanded(!isExpanded);
  };

  let displayedDescription = description;
  let showReadMoreButton = false;

  // Only truncate if it's not HTML and it's longer than the limit
  if (!/<[a-z][\s\S]*>/i.test(description) && description.length > TRUNCATE_LENGTH && !isExpanded) {
    displayedDescription = description.substring(0, TRUNCATE_LENGTH) + "...";
    showReadMoreButton = true;
  }

  // If description contains HTML tags, sanitize and render it
  if (/<[a-z][\s\S]*>/i.test(description)) {
    const sanitizedHtml = React.useMemo(() => 
      htmlSanitizer.sanitizeAlertContent(description), 
      [description]
    );
    
    return (
      <div
        className={richMediaStyles.markdownContainer}
        dangerouslySetInnerHTML={{ __html: sanitizedHtml }}
      />
    );
  }

  const paragraphs = displayedDescription.split("\n\n");

  return (
    <div className={richMediaStyles.markdownContainer}>
      {paragraphs.map((paragraph, index) => {
        // Handle lists
        if (paragraph.includes("\n- ") || paragraph.includes("\n* ")) {
          const [listTitle, ...listItems] = paragraph.split(/\n[-*]\s+/);
          return (
            <div key={`para-${index}`} style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS }}>
              {listTitle.trim() && <Text>{listTitle.trim()}</Text>}
              {listItems.length > 0 && (
                <div style={{ display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS }}>
                  {listItems.map((listItem, itemIndex) => (
                    <div
                      key={`list-item-${itemIndex}`}
                      style={{ display: 'flex', gap: tokens.spacingHorizontalS, alignItems: 'flex-start' }}
                    >
                      <Text>•</Text>
                      <Text>{listItem.trim()}</Text>
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        }

        // Handle bold text
        if (paragraph.includes("**") || paragraph.includes("__")) {
          const parts = paragraph.split(/(\**.*?\**|__.*?__)/g);
          return (
            <Text key={`para-${index}`}>
              {parts.map((part, partIndex) => {
                const isBold = (part.startsWith("**") && part.endsWith("**")) ||
                              (part.startsWith("__") && part.endsWith("__"));
                
                if (isBold) {
                  return (
                    <span
                      key={`part-${partIndex}`}
                      style={{ fontWeight: tokens.fontWeightSemibold }}
                    >
                      {part.slice(2, -2)}
                    </span>
                  );
                }
                return part;
              })}
            </Text>
          );
        }

        // Simple paragraph
        return <Text key={`para-${index}`}>{paragraph}</Text>;
      })}
      {(showReadMoreButton || (description.length > TRUNCATE_LENGTH && isExpanded)) && (
        <Button
          appearance="transparent"
          size="small"
          onClick={toggleExpanded}
          style={{ alignSelf: 'flex-start', marginTop: tokens.spacingVerticalS }}
        >
          {isExpanded ? "Show Less" : "Read More"}
        </Button>
      )}
    </div>
  );
});

export default DescriptionContent;
