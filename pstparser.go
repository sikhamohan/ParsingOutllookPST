package main

import (
	"fmt"
	"os"
	"time"

	pst "github.com/mooijtech/go-pst/v6/pkg"
	"github.com/mooijtech/go-pst/v6/pkg/properties"
	"github.com/rotisserie/eris"
	"golang.org/x/text/encoding"

	charsets "github.com/emersion/go-message/charset"
)

func main() {
	pst.ExtendCharsets(func(name string, enc encoding.Encoding) {
		charsets.RegisterEncoding(name, enc)
	})

	startTime := time.Now()

	fmt.Println("Initializing...")

	//Example usage - path to the pst file
	reader, err := os.Open(".../Outlook.pst")

	if err != nil {
		panic(fmt.Sprintf("Failed to open PST file: %+v\n", err))
	}

	pstFile, err := pst.New(reader)

	if err != nil {
		panic(fmt.Sprintf("Failed to open PST file: %+v\n", err))
	}

	defer func() {
		pstFile.Cleanup()

		if errClosing := reader.Close(); errClosing != nil {
			panic(fmt.Sprintf("Failed to close PST file: %+v\n", err))
		}
	}()

	// Create attachments directory
	if _, err := os.Stat("attachments"); err != nil {
		if err := os.Mkdir("attachments", 0755); err != nil {
			panic(fmt.Sprintf("Failed to create attachments directory: %+v", err))
		}
	}

	// Walk through folders.
	if err := pstFile.WalkFolders(func(folder *pst.Folder) error {

		messageIterator, err := folder.GetMessageIterator()

		if eris.Is(err, pst.ErrMessagesNotFound) {
			// Folder has no messages.
			return nil
		} else if err != nil {
			return err
		}

		// Iterate through messages.
		for messageIterator.Next() {
			message := messageIterator.Value()

			message_props, ok := message.Properties.(*properties.Message)
			if !ok {
				continue
			}
			fmt.Printf("***********************************Email Starts**********************************\n")
			fmt.Printf("Subject:%s\n", message_props.GetSubject())
			deliveryTime := message_props.GetMessageDeliveryTime() / 1e9
			deliveryTimeValue := time.Unix(deliveryTime, 0).UTC()

			// Format the date and time in UTC
			formattedTime := deliveryTimeValue.Format("2006-01-02 15:04:05 MST")
			fmt.Printf("DateandTime%s\n", formattedTime)

			fmt.Printf("Sender:%s\n", message_props.GetSenderEmailAddress())
			fmt.Printf("Receiver:%s\n", message_props.GetReceivedByEmailAddress())
			fmt.Printf("Message:%s\n", message_props.GetBody())
			//fmt.Printf("Body", message_props.String())   // Outputs entire email

			attachmentIterator, err := message.GetAttachmentIterator()

			if eris.Is(err, pst.ErrAttachmentsNotFound) {
				// This message has no attachments.
				continue
			} else if err != nil {
				return err
			}

			// Iterate through attachments.
			for attachmentIterator.Next() {
				attachment := attachmentIterator.Value()

				var attachmentOutputPath string

				if attachment.GetAttachLongFilename() != "" {
					attachmentOutputPath = fmt.Sprintf("attachments/%d-%s", attachment.Identifier, attachment.GetAttachLongFilename())
					fmt.Printf("attachments/%d-%s\n", attachment.Identifier, attachment.GetAttachLongFilename())
				} else {
					attachmentOutputPath = fmt.Sprintf("attachments/UNKNOWN_%d", attachment.Identifier)
					fmt.Printf("attachments/UNKNOWN_%d\n", attachment.Identifier)
				}
				// Saving attachments to attchments folder
				attachmentOutput, err := os.Create(attachmentOutputPath)

				if err != nil {
					return err
				}

				if _, err := attachment.WriteTo(attachmentOutput); err != nil {
					return err
				}

				if err := attachmentOutput.Close(); err != nil {
					return err
				}
			}

			if attachmentIterator.Err() != nil {
				return attachmentIterator.Err()
			}
		}
		return messageIterator.Err()
	}); err != nil {
		panic(fmt.Sprintf("Failed to walk folders: %+v\n", err))
	}

	fmt.Printf("Time: %s\n", time.Since(startTime).String())
}
