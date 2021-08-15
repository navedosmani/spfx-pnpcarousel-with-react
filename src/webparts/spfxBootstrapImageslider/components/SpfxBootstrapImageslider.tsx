import "@pnp/polyfill-ie11";
import * as $ from "jquery";
import "bootstrap/dist/js/bootstrap.bundle.min";
import * as React from "react";
import styles from "./SpfxBootstrapImageslider.module.scss";
import { ISpfxBootstrapImagesliderProps } from "./ISpfxBootstrapImagesliderProps";
import { ISpfxBootstrapImagesliderState } from "./ISpfxBootstrapImagesliderState";
import { Dropdown, DefaultButton, IDropdownOption } from "@fluentui/react";
import { ImageFit } from "@fluentui/react";
import { escape } from "@microsoft/sp-lodash-subset";
import { sp } from "@pnp/sp/presets/all";
import "bootstrap/dist/css/bootstrap.min.css";
import {
  Carousel,
  CarouselButtonsLocation,
  CarouselButtonsDisplay
} from "@pnp/spfx-controls-react/lib/Carousel";

export default class SpfxBootstrapImageslider extends React.Component<
  ISpfxBootstrapImagesliderProps,
  ISpfxBootstrapImagesliderState,
  {}
> {
  constructor(
    props: ISpfxBootstrapImagesliderProps,
    state: ISpfxBootstrapImagesliderState
  ) {
    super(props);
    this.state = {
      images: [],
      statusMessage: "",
      selectedLibray: ""
    };
  }

  public getImagesFromLibrary = async (libraryName: string) => {
    this.setState({ statusMessage: "" });
    if (libraryName) {
      let images: any[] = await sp.web.lists
        .getByTitle(libraryName)
        .items.select(
          "Title, FileRef, EncodedAbsUrl, OData__ExtendedDescription"
        )
        .get();
      //console.log(images);
      //this.setState({ images: images });
      let carouselElements: any[] = [];
      if (images.length > 0) {
        images.map((image) => {
          carouselElements.push({
            imageSrc: image.EncodedAbsUrl,
            title: image.Title,
            description: image.OData__ExtendedDescription,
            url: "https://en.wikipedia.org/wiki/Colosseum",
            showDetailsOnHover: true,
            imageFit: ImageFit.cover
          });
        });
        //console.log(carouselElements);
        this.setState({ images: carouselElements });
      }
    } else {
      this.setState({
        statusMessage: "Please select Image library to load Image Slider"
      });
    }
  };

  public componentDidMount() {
    this.getImagesFromLibrary(this.props.pictureLibraryDropDown);
    //this.getImagesFromLibrary("SliderImages");
  }

  public componentDidUpdate(
    prevProps: ISpfxBootstrapImagesliderProps,
    prevState: ISpfxBootstrapImagesliderState
  ): void {
    if (
      prevProps.pictureLibraryDropDown !== this.props.pictureLibraryDropDown
    ) {
      this.setState({ images: [] });
      this.getImagesFromLibrary(this.props.pictureLibraryDropDown);
    }
  }

  public render(): React.ReactElement<ISpfxBootstrapImagesliderProps> {
    return (
      <React.Fragment>
        {this.state.images.length > 0 ? (
          <div className={styles.spfxBootstrapImageslider}>
            <Carousel
              contentContainerStyles={styles.carouselContent}
              containerButtonsStyles={styles.carouselButtonsContainer}
              buttonsLocation={CarouselButtonsLocation.top}
              buttonsDisplay={CarouselButtonsDisplay.block}
              isInfinite={true}
              element={this.state.images}
            />
          </div>
        ) : (
          <p>No Images found in the selected library</p>
        )}
      </React.Fragment>
    );
  }
}
