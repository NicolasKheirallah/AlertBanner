import * as React from 'react';
import EmojiPickerReact, { EmojiClickData, Theme } from 'emoji-picker-react';
import { Popover, PopoverTrigger, PopoverSurface, Button } from '@fluentui/react-components';
import { Emoji24Regular } from '@fluentui/react-icons';
import styles from './EmojiPicker.module.scss';
import * as strings from 'AlertBannerApplicationCustomizerStrings';

export interface IEmojiPickerProps {
  onEmojiSelect: (emoji: string) => void;
  disabled?: boolean;
  showLabel?: boolean;
}

const EmojiPicker: React.FC<IEmojiPickerProps> = ({
  onEmojiSelect,
  disabled = false,
  showLabel = false
}) => {
  const [isOpen, setIsOpen] = React.useState(false);

  const handleEmojiClick = React.useCallback((emojiData: EmojiClickData) => {
    onEmojiSelect(emojiData.emoji);
    setIsOpen(false);
  }, [onEmojiSelect]);

  return (
    <Popover
      open={isOpen}
      onOpenChange={(e, data) => setIsOpen(data.open)}
      positioning="below-start"
    >
      <PopoverTrigger disableButtonEnhancement>
        <Button
          icon={<Emoji24Regular />}
          disabled={disabled}
          className={styles.emojiButton}
          appearance="subtle"
          title={strings.EmojiPickerButtonTitle}
          aria-label={strings.EmojiPickerButtonTitle}
        >
          {showLabel && strings.EmojiPickerButtonLabel}
        </Button>
      </PopoverTrigger>
      <PopoverSurface className={styles.emojiPickerSurface}>
        <EmojiPickerReact
          onEmojiClick={handleEmojiClick}
          theme={Theme.AUTO}
          searchPlaceHolder={strings.EmojiPickerSearchPlaceholder}
          width={350}
          height={400}
          previewConfig={{
            showPreview: true
          }}
          skinTonesDisabled={false}
          lazyLoadEmojis={true}
        />
      </PopoverSurface>
    </Popover>
  );
};

export default EmojiPicker;
