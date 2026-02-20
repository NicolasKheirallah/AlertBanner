import * as React from "react";
import EmojiPickerReact, { EmojiClickData, Theme } from "emoji-picker-react";
import { Callout, DefaultButton, DirectionalHint } from "@fluentui/react";
import { Emoji24Regular } from "@fluentui/react-icons";
import styles from "./EmojiPicker.module.scss";
import * as strings from "AlertBannerApplicationCustomizerStrings";

const EmojiPicker: React.FC<{
  id?: string;
  onEmojiSelect: (emoji: string) => void;
  disabled?: boolean;
  showLabel?: boolean;
}> = ({ id, onEmojiSelect, disabled = false, showLabel = false }) => {
  const [isOpen, setIsOpen] = React.useState(false);
  const triggerRef = React.useRef<HTMLDivElement | null>(null);

  const handleEmojiClick = React.useCallback(
    (emojiData: EmojiClickData) => {
      onEmojiSelect(emojiData.emoji);
      setIsOpen(false);
    },
    [onEmojiSelect],
  );

  return (
    <>
      <div ref={triggerRef}>
        <DefaultButton
          id={id}
          onRenderIcon={() => <Emoji24Regular />}
          disabled={disabled}
          className={styles.emojiButton}
          title={strings.EmojiPickerButtonTitle}
          ariaLabel={strings.EmojiPickerButtonTitle}
          onClick={() => setIsOpen((prev) => !prev)}
        >
          {showLabel && strings.EmojiPickerButtonLabel}
        </DefaultButton>
      </div>
      {isOpen && triggerRef.current && (
        <Callout
          target={triggerRef.current}
          onDismiss={() => setIsOpen(false)}
          directionalHint={DirectionalHint.bottomLeftEdge}
          gapSpace={8}
          setInitialFocus
          isBeakVisible={false}
          className={styles.emojiPickerSurface}
        >
          <EmojiPickerReact
            onEmojiClick={handleEmojiClick}
            theme={Theme.AUTO}
            searchPlaceHolder={strings.EmojiPickerSearchPlaceholder}
            width={350}
            height={400}
            previewConfig={{
              showPreview: true,
            }}
            skinTonesDisabled={false}
            lazyLoadEmojis={true}
          />
        </Callout>
      )}
    </>
  );
};

export default EmojiPicker;
