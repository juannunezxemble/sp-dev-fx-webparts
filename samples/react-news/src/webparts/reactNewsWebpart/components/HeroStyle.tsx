import * as React from 'react';
import { StylingState, StylingProps } from './StylingPropsState';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { mergeStyles } from 'office-ui-fabric-react/lib/Styling';
import styles from './ReactNewsWebpart.module.scss';
export const iconClass = mergeStyles({
  fontSize: 15,
  height: 15,
  width: 15
});

export default class HeroStyle extends React.Component<StylingProps, StylingState> {

  constructor(props: StylingProps) {
    super(props);
    this.state = {
      News: [],
      RenderedNews: [],
      UpdateCount: 0,
      Next: 4,
      Count: 1,
      Reload: true
    };
  }



  public componentDidMount() {
    var array = [];
    this.props.News.map(Post => {
        array.push(Post);
    });
    this.setState({ RenderedNews: array, UpdateCount: 0 });
  }
  public componentDidUpdate(prevProps: StylingProps) {
    var array = [];
    if (prevProps.News !== this.props.News) {

      this.props.News.map(Post => {
          array.push(Post);
      });
      this.setState({ RenderedNews: array, UpdateCount: 0 });
      return true;
    }
    else if (this.props.News.length > 0 && this.props.News.length > this.state.RenderedNews.length) {
      this.props.News.map(Post => {
          array.push(Post);
      });
      this.setState({ RenderedNews: array, UpdateCount: this.state.UpdateCount + 1 });
      return true;
    }
  }
  public render(): React.ReactElement<StylingProps> {
    var i = 0;

    return (<div className={styles.SingleStyle}>
      <div  className={styles.SingleStyleContainer} >
        <div>{this.state.RenderedNews.map(Post => {
          return <div className={styles.NewsContainer} style={{ boxShadow: 'rgb(0 0 0 / 16%) 0px 1px 4px, rgb(0 0 0 / 10%) 0px 0px 1px' }}>
            <div className={styles.ImgContainer}>
              <img src={Post.Thumbnail} className={styles.Image}></img></div>
            <div className={styles.NewsBody}>
              <div className={styles.TitleContainer}>
                <a className={styles.TitleStyling} href={Post.Url}>{Post.Title}</a>
              </div>
              <div className={styles.DescriptionContainer}>{Post.Description}</div>
              <div className={styles.IconContainer}>
                <Icon className={iconClass} iconName="Like"></Icon>
                <label className={styles.IconLabelStyling}>
                  {Post.Likes}
                </label>
                <Icon style={{ marginLeft: "10px" }} className={iconClass} iconName="Comment"></Icon>
                <label className={styles.IconLabelStyling}>
                  {Post.Comments}
                </label>
              </div>
              <div className={styles.AuthorContainer}>
                {this.props.AuthorToggle ? <></> : Post.Author} Created {Post.Created}
              </div>
            </div>
          </div>;
        })}</div>
      
      </div>
    </div>
    );
  }
}