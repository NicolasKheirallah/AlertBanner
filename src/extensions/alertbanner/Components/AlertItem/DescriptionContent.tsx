import * as React from "react";
import { Button, Text } from "@fluentui/react-components";
import { htmlSanitizer } from "../Utils/HtmlSanitizer";
import styles from "./AlertItem.module.scss";
import { useLocalization } from "../Hooks/useLocalization";

interface IDescriptionContentProps {
  description: string;
}

const DescriptionContent: React.FC<IDescriptionContentProps> = React.memo(({ description }) => {
  const [isExpanded, setIsExpanded] = React.useState(false);
  const TRUNCATE_LENGTH = 200;
  const { getString } = useLocalization();

  const containsHtml = React.useMemo(() => /<[a-z][\s\S]*>/i.test(description), [description]);
  const sanitizedHtml = React.useMemo(() => {
    if (!containsHtml) {
      return '';
    }

    return htmlSanitizer.sanitizeAlertContent(description);
  }, [containsHtml, description]);

  const toggleExpanded = () => {
    setIsExpanded(!isExpanded);
  };

  let displayedDescription = description;
  let showReadMoreButton = false;

  if (!containsHtml && description.length > TRUNCATE_LENGTH && !isExpanded) {
    displayedDescription = description.substring(0, TRUNCATE_LENGTH) + "...";
    showReadMoreButton = true;
  }

  if (containsHtml) {
    return (
      <div
        className={styles.descriptionListContainer}
        dangerouslySetInnerHTML={{ __html: sanitizedHtml }}
      />
    );
  }

  const paragraphs = displayedDescription.split("\n\n");

  return (
    <div className={styles.descriptionListContainer}>
      {paragraphs.map((paragraph, index) => {
        if (paragraph.includes("\n- ") || paragraph.includes("\n* ")) {
          const [listTitle, ...listItems] = paragraph.split(/\n[-*]\s+/);
          return (
            <div key={`para-${index}`} className={styles.descriptionParagraph}>
              {listTitle.trim() && <Text>{listTitle.trim()}</Text>}
              {listItems.length > 0 && (
                <div className={styles.descriptionListContainer}>
                  {listItems.map((listItem, itemIndex) => (
                    <div
                      key={`list-item-${itemIndex}`}
                      className={styles.descriptionListItem}
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
                      className={styles.descriptionBoldText}
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

        return <Text key={`para-${index}`}>{paragraph}</Text>;
      })}
      {(showReadMoreButton || (description.length > TRUNCATE_LENGTH && isExpanded)) && (
        <Button
          appearance="transparent"
          size="small"
          onClick={toggleExpanded}
          className={styles.descriptionToggleButton}
        >
          {isExpanded ? getString('ShowLess') : getString('ShowMore')}
        </Button>
      )}
    </div>
  );
});

DescriptionContent.displayName = 'DescriptionContent';

export default DescriptionContent;
