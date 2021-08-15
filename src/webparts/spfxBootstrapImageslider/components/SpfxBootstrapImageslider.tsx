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
import Carousel from "react-bootstrap/Carousel";

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
      selectedLibray: "",
      index: 0
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
      console.log(images);
      //this.setState({ images: images });
      if (images.length > 0) {
        //console.log(carouselElements);
        this.setState({ images: images });
      }
    } else {
      this.setState({
        statusMessage: "Please select Image library to load Image Slider"
      });
    }
  };

  public handleSelect = (selectedIndex, e) => {
    this.setState({ index: selectedIndex });
  };

  componentDidMount() {
    this.getImagesFromLibrary(this.props.pictureLibraryDropDown);
    //this.getImagesFromLibrary("SliderImages");
  }

  componentDidUpdate(
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
        <div className={styles.spfxBootstrapImageslider}>
          {this.state.images.length > 0 ? (
            <Carousel
              fade
              variant="dark"
              activeIndex={this.state.index}
              interval={2000}
              onSelect={this.handleSelect}
            >
              {this.state.images.map((image) => {
                return (
                  <Carousel.Item
                    className={styles.reactBootstrapCarouselContent}
                  >
                    <img
                      className="d-block w-100"
                      src={image.EncodedAbsUrl}
                      alt={image.Title}
                    />
                    <Carousel.Caption>
                      <h3>{image.Title}</h3>
                      <p>{image.OData__ExtendedDescription}</p>
                    </Carousel.Caption>
                  </Carousel.Item>
                );
              })}
            </Carousel>
          ) : (
            <p>
              No images avaiable in the selected library. Please select another
              library.
            </p>
          )}
        </div>
      </React.Fragment>
    );
  }
}
